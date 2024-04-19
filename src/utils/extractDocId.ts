export const extractDocId = (url: string): string | null => {
  const match = /\/d\/(.*?)(\/|$)/.exec(url);
  return match ? match[1] : null;
};
