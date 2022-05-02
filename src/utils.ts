export const chunkArray = <T>(array: T[], chunkSize: number): T[][] => {
  return Array.from({ length: Math.ceil(array.length / chunkSize) }, (_, index) =>
    array.slice(index * chunkSize, (index + 1) * chunkSize)
  );
};

export const buildWhereInStringValue = (arr: string[]): string => {
  if (!arr || !arr.length) {
    return '';
  }
  return arr
    .map((val) => {
      return `'${val}'`;
    })
    .join(',');
};

export const getMetadataPropAsArray = <T>(prop: string, metadata: unknown): T[] => {
  if (Array.isArray(metadata[prop])) {
    return metadata[prop] as T[];
  } else if (!Array.isArray(metadata[prop]) && typeof metadata[prop] === 'object' && metadata[prop] !== null) {
    return [metadata[prop]] as T[];
  } else {
    return metadata[prop] as T[];
  }
};
