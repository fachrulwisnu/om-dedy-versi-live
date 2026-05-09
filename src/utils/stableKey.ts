export const createStableKey = (
  prefix: string,
  id?: string | number | null,
  index?: number
) => {
  const safeId =
    id !== undefined &&
    id !== null &&
    String(id).trim() !== ''
      ? String(id)
      : `fallback-${index ?? 'unknown'}`;

  return `${prefix}-${safeId}`;
};
