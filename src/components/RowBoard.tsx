import React, { useMemo, useRef, useState, useCallback, useEffect } from "react";
import {
  DndContext,
  DragOverlay,
  PointerSensor,
  closestCenter,
  pointerWithin,
  useDroppable,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import type { DragEndEvent, DragOverEvent, DragStartEvent } from "@dnd-kit/core";
import {
  SortableContext,
  arrayMove,
  rectSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import type { ImageData } from "../types";

type Props = {
  images: ImageData[];
  setImages: React.Dispatch<React.SetStateAction<ImageData[]>>;
  rows?: number;
  setActiveCropImageId?: (id: string) => void;
  onUnassignImage?: (id: string) => void;
};

// ✅ 共通カードUI（SortableItem と DragOverlay で使い回す）
function ItemContent({ img, setImages, setActiveCropImageId, onUnassignImage }: { img: ImageData; setImages: React.Dispatch<React.SetStateAction<ImageData[]>>; setActiveCropImageId?: (id: string) => void; onUnassignImage?: (id: string) => void }) {
  const handleDelete = () => {
    setImages(prev => prev.filter(i => i.id !== img.id));
  };
  const handleUnassign = () => {
    if (onUnassignImage) {
      onUnassignImage(img.id);
    } else {
      setImages(prev => prev.map(i => i.id === img.id ? { ...i, row: 0 } : i));
      if (setActiveCropImageId) setActiveCropImageId(img.id);
    }
  };
  return (
    <div
      style={{
        border: "1px solid #eee",
        borderRadius: 10,
        background: "#fafafa",
        overflow: "hidden",
        width: 120,
        userSelect: "none",
      }}
    >
      <div style={{ padding: 8 }}>
        <div
          style={{
            width: "100%",
            height: 80,
            borderRadius: 8,
            background: "#eaeaea",
            overflow: "hidden",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <img
            src={img.dataUrl}
            alt={img.name}
            style={{
              maxWidth: "100%",
              maxHeight: "100%",
              transform: `rotate(${img.rotation ?? 0}deg)`,
              pointerEvents: "none",
            }}
          />
        </div>
        <div
          style={{
            fontSize: 12,
            marginTop: 6,
            whiteSpace: "nowrap",
            overflow: "hidden",
            textOverflow: "ellipsis",
          }}
        >
          {/* ファイル名表示を削除（img.name） */}
        </div>
        <div style={{ display: 'flex', gap: 4, marginTop: 8 }}>
          <button
            onClick={handleDelete}
            style={{
              flex: 1,
              background: '#ffe4e6',
              color: '#be123c',
              border: '1px solid #fca5a5',
              borderRadius: 6,
              padding: '2px 0',
              fontSize: 12,
              cursor: 'pointer',
            }}
            title="削除"
          >削除</button>
          {img.row !== 0 && (
            <button
              onClick={handleUnassign}
              style={{
                flex: 1,
                background: '#f1f5f9',
                color: '#334155',
                border: '1px solid #cbd5e1',
                borderRadius: 6,
                padding: '2px 0',
                fontSize: 12,
                cursor: 'pointer',
              }}
              title="戻す"
            >戻す</button>
          )}
        </div>
      </div>
    </div>
  );
}

// ✅ ドラッグ可能な個別アイテム
function SortableItem({ img, setImages, setActiveCropImageId, onUnassignImage }: { img: ImageData; setImages: React.Dispatch<React.SetStateAction<ImageData[]>>; setActiveCropImageId?: (id: string) => void; onUnassignImage?: (id: string) => void }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } =
    useSortable({ id: img.id });

  return (
    <div
      ref={setNodeRef}
      style={{
        transform: CSS.Transform.toString(transform),
        transition,
        opacity: isDragging ? 0.3 : 1,
        cursor: "grab",
      }}
      {...attributes}
      {...listeners}
    >
      <ItemContent img={img} setImages={setImages} setActiveCropImageId={setActiveCropImageId} onUnassignImage={onUnassignImage} />
    </div>
  );
}

// ✅ 段落コンテナ（droppable + sortable context）
function RowContainer({ row, images, setImages, setActiveCropImageId, onUnassignImage }: { row: number; images: ImageData[]; setImages: React.Dispatch<React.SetStateAction<ImageData[]>>; setActiveCropImageId?: (id: string) => void; onUnassignImage?: (id: string) => void }) {
  const ids = images.map((img) => img.id);
  const { setNodeRef, isOver } = useDroppable({ id: `row-${row}` });

  return (
    <div
      ref={setNodeRef}
      style={{
        border: isOver ? "2px solid #3b82f6" : "1px solid #ddd",
        borderRadius: 12,
        padding: 12,
        background: isOver ? "#eff6ff" : "#fff",
        transition: "all 0.12s ease",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
        <div style={{ fontWeight: 700 }}>段落 {row}</div>
        <div style={{ fontSize: 12, opacity: 0.7 }}>{images.length} 枚</div>
      </div>
      <SortableContext items={ids} strategy={rectSortingStrategy}>
        <div
          style={{
            display: "flex",
            gap: 10,
            flexWrap: "wrap",
            minHeight: 92,
          }}
        >
          {images.map((img) => (
            <SortableItem key={img.id} img={img} setImages={setImages} setActiveCropImageId={setActiveCropImageId} onUnassignImage={onUnassignImage} />
          ))}
        </div>
      </SortableContext>
    </div>
  );
}

export default function RowBoard({ images, setImages, rows = 4, setActiveCropImageId, onUnassignImage }: Props) {
  const [activeId, setActiveId] = useState<string | null>(null);
  // ✅ ドラッグ中アイテムの現在行（stale closure 回避のため ref で管理）
  const activeItemRowRef = useRef<number | null>(null);
  // ✅ 空段落の追跡（collisionDetection が参照、byRow の useMemo 内で更新）
  const emptyRowsRef = useRef<Set<number>>(new Set());

  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } })
  );

  // ✅ 衝突検出：アイテム優先戦略
  //   アイテム上          → アイテム ID を返す（SortableContext が transform を計算 → スライドアニメ）
  //   コンテナ内の隙間    → 最近傍アイテム ID（closestCenter）
  //   空コンテナ上        → コンテナ ID（段落間移動のトリガー）
  //   コンテナ外          → closestCenter フォールバック
  const collisionDetection = useCallback(
    (args: Parameters<typeof closestCenter>[0]) => {
      const pointerHits = pointerWithin(args);
      const itemHits = pointerHits.filter(({ id }) => !String(id).startsWith("row-"));

      if (itemHits.length > 0) {
        // アイテム上：アイテムのみを対象に closestCenter
        return closestCenter({
          ...args,
          droppableContainers: args.droppableContainers.filter(
            (c) => !String(c.id).startsWith("row-")
          ),
        });
      }

      const containerHits = pointerHits.filter(({ id }) => String(id).startsWith("row-"));
      if (containerHits.length > 0) {
        const rowNum = parseInt(String(containerHits[0].id).replace("row-", ""), 10);
        if (emptyRowsRef.current.has(rowNum)) {
          // 空コンテナ：コンテナ ID を返して段落間移動を検出
          return containerHits;
        }
        // 非空コンテナの隙間：最近傍アイテム ID を返す
        const itemContainers = args.droppableContainers.filter(
          (c) => !String(c.id).startsWith("row-")
        );
        if (itemContainers.length > 0) {
          return closestCenter({ ...args, droppableContainers: itemContainers });
        }
      }

      return closestCenter(args);
    },
    [] // emptyRowsRef は ref なので deps 不要
  );

  const byRow = useMemo(() => {
    const map = new Map<number, ImageData[]>();
    for (let r = 1; r <= rows; r++) map.set(r, []);
    for (const img of images) {
      const r = img.row ?? 0;
      if (r >= 1 && r <= rows) {
        map.get(r)!.push(img);
      }
    }
    return map;
  }, [images, rows]);

  // ✅ 空段落を副作用で更新（useMemo 内の副作用を排除）
  useEffect(() => {
    const emptyRows = new Set<number>();
    for (let r = 1; r <= rows; r++) {
      if ((byRow.get(r) ?? []).length === 0) emptyRows.add(r);
    }
    emptyRowsRef.current = emptyRows;
  }, [byRow, rows]);

  const activeImage = useMemo(
    () => (activeId ? images.find((img) => img.id === activeId) ?? null : null),
    [activeId, images]
  );

  const handleDragStart = useCallback(
    (event: DragStartEvent) => {
      const id = String(event.active.id);
      setActiveId(id);
      const img = images.find((x) => x.id === id);
      activeItemRowRef.current = img?.row ?? null;
    },
    [images]
  );

  const handleDragOver = useCallback(
    (event: DragOverEvent) => {
      const { over } = event;
      if (!over) return;

      const overId = String(over.id);

      // ターゲット段落を特定
      let targetRow: number | null = null;
      if (overId.startsWith("row-")) {
        targetRow = parseInt(overId.replace("row-", ""), 10);
      } else {
        // 別アイテムの上にいる場合、そのアイテムの段落を使う
        const overImg = images.find((img) => img.id === overId);
        if (!overImg) return;
        targetRow = overImg.row ?? null;
      }

      if (!targetRow) return;

      // 同じ段落なら何もしない（onDragEnd の arrayMove に任せる）
      if (activeItemRowRef.current === targetRow) return;

      // ドロップ候補だけ記録する（state は更新しない）
      activeItemRowRef.current = targetRow;
    },
    [images]
  );

  const handleDragEnd = useCallback(
    (event: DragEndEvent) => {
      const { active, over } = event;
      setActiveId(null);
      if (!over) {
        activeItemRowRef.current = null;
        return;
      }

      const activeIdStr = String(active.id);
      const overId = String(over.id);

      setImages((prev) => {
        const activeIdx = prev.findIndex((img) => img.id === activeIdStr);
        if (activeIdx === -1) return prev;

        // 段落間移動は onDragEnd でのみ確定
        const activeRow = prev[activeIdx].row ?? null;
        let targetRow: number | null = activeItemRowRef.current;
        if (overId.startsWith("row-")) {
          targetRow = parseInt(overId.replace("row-", ""), 10);
        } else {
          const overImg = prev.find((img) => img.id === overId);
          if (overImg) targetRow = overImg.row ?? null;
        }

        if (targetRow && activeRow !== targetRow) {
          const result = [...prev];
          result[activeIdx] = { ...result[activeIdx], row: targetRow };
          return result;
        }

        // コンテナ上へのドロップ（空エリアへのドロップ）は並び替え不要
        if (overId.startsWith("row-") || activeIdStr === overId) return prev;

        // 同段落内の並び替え
        const activeImg = prev[activeIdx];
        const overImg = prev.find((img) => img.id === overId);
        if (!overImg) return prev;
        if ((activeImg.row ?? 0) !== (overImg.row ?? 0)) return prev;

        const row = activeImg.row ?? 0;
        const rowItems = prev.filter((img) => (img.row ?? 0) === row);
        const others = prev.filter((img) => (img.row ?? 0) !== row);
        const ids = rowItems.map((img) => img.id);
        const oldIndex = ids.indexOf(activeIdStr);
        const newIndex = ids.indexOf(overId);
        if (oldIndex === -1 || newIndex === -1) return prev;

        return [...others, ...arrayMove(rowItems, oldIndex, newIndex)];
      });
      activeItemRowRef.current = null;
    },
    [setImages]
  );

  const handleDragCancel = useCallback(() => {
    setActiveId(null);
    activeItemRowRef.current = null;
  }, []);

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={collisionDetection}
      onDragStart={handleDragStart}
      onDragOver={handleDragOver}
      onDragEnd={handleDragEnd}
      onDragCancel={handleDragCancel}
    >
      <div style={{ display: "grid", gap: 12 }}>
        {Array.from({ length: rows }, (_, i) => i + 1).map((row) => (
          <RowContainer key={row} row={row} images={byRow.get(row) ?? []} setImages={setImages} setActiveCropImageId={setActiveCropImageId} onUnassignImage={onUnassignImage} />
        ))}
      </div>
      <DragOverlay dropAnimation={null}>
        {activeImage ? <ItemContent img={activeImage} setImages={setImages} setActiveCropImageId={setActiveCropImageId} /> : null}
      </DragOverlay>
    </DndContext>
  );
}
