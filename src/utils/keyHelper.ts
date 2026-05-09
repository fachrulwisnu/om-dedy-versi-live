export const getSafeKey = (item: any, index: number, prefix: string) => {
  const baseId = item?.custom_id || item?.id || item?._id || 'static';
  return `${prefix}-${baseId}-${index}`;
};
