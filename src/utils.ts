import { Connection } from '@salesforce/core';

export const chunkArray = <T>(array: T[], chunkSize: number): T[][] => {
  const chunkedArr: T[][] = [];
  let index = 0;
  while (index < array.length) {
    chunkedArr.push(array.slice(index, chunkSize + index));
    index += chunkSize;
  }
  return chunkedArr;
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
  } else if (typeof metadata[prop] === 'string') {
    return [metadata[prop]] as T[];
  } else {
    return metadata[prop] as T[];
  }
};

export const getMetadataAsArray = async <T>(
  conn: Connection,
  metadataType: string,
  fullNames: string[]
): Promise<T[]> => {
  const metadataResult = (await conn.metadata.read(metadataType, fullNames)) as unknown;
  let metadataRecords: T[];
  if (Array.isArray(metadataResult)) {
    metadataRecords = metadataResult as T[];
  } else {
    metadataRecords = [metadataResult as T];
  }
  return metadataRecords;
};