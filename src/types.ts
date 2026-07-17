export type Marketplace = 'shopee' | 'tokopedia';

export type TabType = 'mismatch' | 'match' | 'only1' | 'only2' | 'discontinued' | 'shortSku';

export type Data1Source = 'auto_balist' | 'auto_gomall' | 'manual';

export interface MismatchItem {
  sku: string;
  name: string;
  stock1: number;
  stockAio: number;
}

export interface MatchItem {
  sku: string;
  name: string;
  stock1: number;
  stockAio: number;
}

export interface Only1Item {
  sku: string;
  stock1: number;
}

export interface Only2Item {
  sku: string;
  name: string;
  stockAio: number;
  category: string;
  brand: string;
}

export interface DiscontinuedItem {
  sku: string;
  name: string;
  stock1: number;
  stock3: number;
}

export interface ShortSkuItem {
  sku: string;
  suggestedSku: string;
  name: string;
  source: string;
  stock: number;
  existsInAio?: boolean;
}

export interface ComparisonResults {
  mismatch: MismatchItem[];
  match: MatchItem[];
  only1: Only1Item[];
  only2: Only2Item[];
  discontinued: DiscontinuedItem[];
  shortSku: ShortSkuItem[];
}
