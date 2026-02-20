import React, { useMemo, useRef, useState, useCallback } from "react";
import { Reorder } from "framer-motion";
import type { ImageData } from "../types";

type Props = {
  images: ImageData[];
  setImages: React.Dispatch<React.SetStateAction<ImageData[]>>;
  rows?: number; // 1..N
};

export default function RowBoard({ images, setImages, rows = 4 }: Props) {
  const rowRefs = useRef<Record<number, HTMLDivElement | null>>({});
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const [overRow, setOverRow] = useState<number | null>(null);

  // ✅ 重要：並び順は sort しない（Reorderの結果が壊れる）
  const byRow = useMemo(() => {
    const map = new Map<number, ImageData[]>();
    for (let r = 0; r <= rows; r++) map.set(r, []);
    for (const img of images) {
      const r = img.row ?? 0;
      if (!map.has(r)) map.set(r, []);
      map.get(r)!.push(img);
    }
    return map;
  }, [images, rows]);

  // ✅ 段落内：id順で並び替え
  const reorderWithinRowByIds = useCallback(
    (row: number, orderedIds: string[]) => {
      setImages((prev) => {
        const rowItems = prev.filter((p) => (p.row ?? 0) === row);
        const others = prev.filter((p) => (p.row ?? 0) !== row);

        const byId = new Map(rowItems.map((x) => [x.id, x]));
        const nextRow: ImageData[] = [];

        for (const id of orderedIds) {
          const item = byId.get(id);
          if (item) nextRow.push(item);
        }
        // 念のため漏れがあれば末尾に
        for (const item of rowItems) {
          if (!orderedIds.includes(item.id)) nextRow.push(item);
        }

        return [...others, ...nextRow];
      });
    },
    [setImages]
  );

  // ✅ 別段落へ移動
  const moveToRow = useCallback(
    (imageId: string, toRow: number) => {
      setImages((prev) =>
        prev.map((img) =>
          img.id === imageId ? { ...img, row: toRow, orderConfirmed: false } : img
        )
      );
    },
    [setImages]
  );

  return (
    <div style={{ display: "grid", gap: 12 }}>
      {Array.from({ length: rows + 1 }, (_, i) => i).map((row) => {
        const list = byRow.get(row) ?? [];
        const title = row === 0 ? "未割当" : `段落 ${row}`;
        const orderedIds = list.map((x) => x.id);

        return (
          <div
            key={row}
            ref={(el) => {
              rowRefs.current[row] = el;
            }}
            // ✅ 段落移動（HTML Drag&Drop）用
            onDragEnter={() => {
              if (!draggingId) return;
              setOverRow(row);
            }}
            onDragOver={(e) => {
              // ✅ ここが超重要：drop可能にするため必ず preventDefault
              if (!draggingId) return;
              e.preventDefault();
              setOverRow(row);
            }}
            onDragLeave={() => {
              if (overRow === row) setOverRow(null);
            }}
            onDrop={(e) => {
              e.preventDefault();
              const id = e.dataTransfer.getData("text/plain");
              if (id) moveToRow(id, row);
              setDraggingId(null);
              setOverRow(null);
            }}
            style={{
              border: overRow === row && draggingId ? "2px solid #3b82f6" : "1px solid #ddd",
              borderRadius: 12,
              padding: 12,
              background: overRow === row && draggingId ? "#eff6ff" : "#fff",
              transition: "all 0.12s ease",
            }}
          >
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <div style={{ fontWeight: 700 }}>{title}</div>
              <div style={{ fontSize: 12, opacity: 0.7 }}>{list.length} 枚</div>
            </div>

            {/* ✅ 段落内並び替え（Reorder） */}
            <Reorder.Group
              axis="x"
              values={orderedIds}
              onReorder={(newIds) => reorderWithinRowByIds(row, newIds)}
              style={{
                display: "flex",
                gap: 10,
                paddingTop: 10,
                flexWrap: "wrap",
                minHeight: 92,
              }}
            >
              {list.map((img) => (
                <Reorder.Item
                  key={img.id}
                  value={img.id}
                  style={{ width: 120, userSelect: "none" }}
                  whileDrag={{ scale: 1.03 }}
                >
                  <div
                    style={{
                      border: "1px solid #eee",
                      borderRadius: 10,
                      background: "#fafafa",
                      overflow: "hidden",
                    }}
                    title="カード本体：同じ段落内の並び替え"
                  >
                    {/* ✅ 段落移動用ハンドル（ここだけ draggable） */}
                    <div
                      draggable
                      onPointerDownCapture={(e) => e.stopPropagation()} // ✅ Reorderに触らせない
                      onMouseDownCapture={(e) => e.stopPropagation()}   // ✅ 同上（Windows対策）
                      onDragStart={(e) => {
                        e.stopPropagation(); // ✅ これも効く
                        setDraggingId(img.id);
                        e.dataTransfer.setData("text/plain", img.id);
                        e.dataTransfer.effectAllowed = "move";
                      }}
                      onDragEnd={() => {
                        setDraggingId(null);
                        setOverRow(null);
                      }}
                      style={{
                        padding: "6px 8px",
                        fontSize: 12,
                        fontWeight: 700,
                        background: "#e5e7eb",
                        cursor: "grab",
                      }}
                      title="ここをドラッグすると別段落へ移動"
                    >
                      段落移動
                    </div>

                    {/* ✅ 画像本体（draggable無し＝Reorderが効く） */}
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
                        {img.name}
                      </div>
                    </div>
                  </div>
                </Reorder.Item>
              ))}
            </Reorder.Group>
          </div>
        );
      })}
    </div>
  );
}