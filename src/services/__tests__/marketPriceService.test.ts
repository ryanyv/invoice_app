import { describe, it, expect, beforeEach, vi } from 'vitest';
import {
  MarketPriceService,
  BASELINE_MARKET_ITEMS,
  BASELINE_MARKET_SNAPSHOT,
  MARKET_STORAGE_KEY,
} from '../marketPriceService';
import { MarketPriceSnapshot } from '../../types';
import { StorageService } from '../storage';

// Setup Mock LocalStorage for Node test environment
const store = new Map<string, string>();
const mockLocalStorage = {
  getItem: (key: string) => store.get(key) ?? null,
  setItem: (key: string, value: string) => store.set(key, String(value)),
  removeItem: (key: string) => store.delete(key),
  clear: () => store.clear(),
};
(globalThis as any).localStorage = mockLocalStorage;

describe('MarketPriceService (Iranian Polymer Commodity & IME Prices)', () => {
  beforeEach(() => {
    mockLocalStorage.clear();
    vi.restoreAllMocks();
  });

  describe('1. Baseline Data & Structure', () => {
    it('contains all standard pipe grades and carbon black masterbatch', () => {
      expect(BASELINE_MARKET_ITEMS.length).toBeGreaterThanOrEqual(4);

      const pe100 = BASELINE_MARKET_ITEMS.find((i) => i.grade === 'PE100');
      const pe80 = BASELINE_MARKET_ITEMS.find((i) => i.grade === 'PE80');
      const pe63 = BASELINE_MARKET_ITEMS.find((i) => i.grade === 'PE63');
      const mb = BASELINE_MARKET_ITEMS.find((i) => i.grade === 'Masterbatch');

      expect(pe100).toBeDefined();
      expect(pe80).toBeDefined();
      expect(pe63).toBeDefined();
      expect(mb).toBeDefined();

      expect(pe100!.pricePerKgToman).toBeGreaterThan(50000);
      expect(pe80!.pricePerKgToman).toBeGreaterThan(40000);
      expect(pe63!.pricePerKgToman).toBeGreaterThan(40000);
      expect(mb!.pricePerKgToman).toBeGreaterThan(80000);
    });

    it('has valid Persian names, official market descriptions, and producers', () => {
      for (const item of BASELINE_MARKET_ITEMS) {
        expect(item.name).toBeTruthy();
        expect(item.producer).toBeTruthy();
        expect(item.market).toBeTruthy();
        expect(item.pricePerKgToman).toBeTypeOf('number');
      }
    });
  });

  describe('2. Local Caching & Offline Fallback', () => {
    it('returns BASELINE_MARKET_SNAPSHOT when localStorage is empty', () => {
      const stored = MarketPriceService.getStoredMarketPrices();
      expect(stored).toBeDefined();
      expect(stored.items.length).toBe(BASELINE_MARKET_ITEMS.length);
      expect(stored.isOnline).toBe(false);
    });

    it('saves and retrieves updated market snapshots from cache', () => {
      const customSnapshot: MarketPriceSnapshot = {
        source: 'آزمون نرخ مصوب',
        lastUpdated: '1405/01/01',
        isOnline: true,
        items: [
          {
            id: 'test_pe100',
            grade: 'PE100',
            name: 'پلی‌اتیلن تست',
            producer: 'پتروشیمی جم',
            code: 'CRP100N',
            pricePerKgToman: 89000,
            market: 'بورس کالای ایران',
            publishDate: '1405/01/01',
          },
        ],
      };

      MarketPriceService.saveMarketPrices(customSnapshot);
      const retrieved = MarketPriceService.getStoredMarketPrices();

      expect(retrieved.source).toBe('آزمون نرخ مصوب');
      expect(retrieved.items[0].pricePerKgToman).toBe(89000);
      expect(retrieved.isOnline).toBe(true);
    });

    it('recovers gracefully to baseline if stored JSON is corrupted', () => {
      localStorage.setItem(MARKET_STORAGE_KEY, '{ invalid_json :::');
      const retrieved = MarketPriceService.getStoredMarketPrices();
      expect(retrieved).toEqual(BASELINE_MARKET_SNAPSHOT);
    });
  });

  describe('3. Offline Resilience & Live Fetching', () => {
    it('returns cached data marked as offline when network request throws or fails', async () => {
      // Simulate offline / network disconnection
      global.fetch = vi.fn().mockRejectedValue(new Error('Failed to fetch / Offline'));

      const result = await MarketPriceService.fetchLiveMarketPrices();
      expect(result).toBeDefined();
      expect(result.isOnline).toBe(false);
      expect(result.statusMessage).toContain('آفلاین');
      expect(result.items.length).toBeGreaterThanOrEqual(4);
    });

    it('uses Electron IPC when running inside Electron desktop container', async () => {
      const mockIpcSnapshot: MarketPriceSnapshot = {
        source: 'بورس کالای ایران - الکترون',
        lastUpdated: '1405/07/25',
        isOnline: true,
        items: BASELINE_MARKET_ITEMS.map((item) => ({ ...item, pricePerKgToman: item.pricePerKgToman + 500 })),
      };

      (globalThis as any).window = (globalThis as any).window || {};
      (globalThis as any).window.electronAPI = {
        fetchMarketPrices: vi.fn().mockResolvedValue({
          success: true,
          isOnline: true,
          snapshot: mockIpcSnapshot,
        }),
      };

      const result = await MarketPriceService.fetchLiveMarketPrices();
      expect(result.isOnline).toBe(true);
      expect(result.source).toBe('بورس کالای ایران - الکترون');
      expect(result.items[0].pricePerKgToman).toBe(BASELINE_MARKET_ITEMS[0].pricePerKgToman + 500);

      delete (globalThis as any).window.electronAPI;
    });
  });

  describe('4. Applying Official Market Prices to Company Settings', () => {
    it('applyMarketPriceToGrade updates single grade price in settings', () => {
      const updated = MarketPriceService.applyMarketPriceToGrade('PE100', 91500);
      expect(updated.basePipePrices.PE100).toBe(91500);

      const fromStorage = StorageService.getSettings();
      expect(fromStorage.basePipePrices.PE100).toBe(91500);
    });

    it('applyAllMarketPrices updates PE100, PE80, and PE63 prices in settings', () => {
      const snapshot: MarketPriceSnapshot = {
        source: 'بورس کالا',
        lastUpdated: '1405/01/01',
        isOnline: true,
        items: [
          {
            id: 'ime_pe100_black',
            grade: 'PE100',
            name: 'PE100',
            producer: 'جم',
            code: 'CRP100N',
            pricePerKgToman: 88800,
            market: 'IME',
            publishDate: '1405/01/01',
          },
          {
            id: 'ime_pe80_pipe',
            grade: 'PE80',
            name: 'PE80',
            producer: 'امیرکبیر',
            code: 'PE80',
            pricePerKgToman: 79900,
            market: 'IME',
            publishDate: '1405/01/01',
          },
          {
            id: 'ime_pe63_pipe',
            grade: 'PE63',
            name: 'PE63',
            producer: 'باختر',
            code: 'PE63',
            pricePerKgToman: 71100,
            market: 'IME',
            publishDate: '1405/01/01',
          },
        ],
      };

      const updated = MarketPriceService.applyAllMarketPrices(snapshot);
      expect(updated.basePipePrices.PE100).toBe(88800);
      expect(updated.basePipePrices.PE80).toBe(79900);
      expect(updated.basePipePrices.PE63).toBe(71100);

      const fromStorage = StorageService.getSettings();
      expect(fromStorage.basePipePrices.PE100).toBe(88800);
      expect(fromStorage.basePipePrices.PE80).toBe(79900);
      expect(fromStorage.basePipePrices.PE63).toBe(71100);
    });
  });
});
