const stampModules = import.meta.glob('./assets/stamps/*.{png,PNG}', {
  eager: true,
  import: 'default',
  query: '?inline'
}) as Record<string, string>;

export function sanitizeStampKey(rawName: string): string {
  return String(rawName ?? '')
    .normalize('NFKC')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[\\/:*?"<>|]/g, '')
    .trim();
}

const stampUrlMap = new Map<string, string>();

Object.entries(stampModules).forEach(([filePath, url]) => {
  const fileName = filePath.split('/').pop() ?? '';
  const baseName = fileName.replace(/\.[^/.]+$/, '');
  const key = sanitizeStampKey(baseName);
  if (!key) return;
  stampUrlMap.set(key, url);
});

export function getStampUrlByVetName(vetName: string): string | null {
  const key = sanitizeStampKey(vetName);
  if (!key) return null;
  return stampUrlMap.get(key) ?? null;
}
