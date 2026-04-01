import { promisify } from 'util';
import * as crypto from 'crypto';

export interface CachedFormFill {
  formHash: string;
  fillActions: AiCachedFillAction[];
}

export interface AiCachedFillAction {
  selector: string;
  value: string;
  inputType: 'text' | 'select' | 'checkbox' | 'radio' | 'click';
}

/**
 * Redis-backed cache for AI form fill results.
 *
 * Key difference from cog-marketo's CachingClientWrapper:
 *   - cog-marketo: 55s TTL, per-scenario-run scope — deduplicates within a single run
 *   - AiFormCache:  8-day TTL, per-org+URL scope   — persists across runs until form structure changes
 *
 * Cache key: WebCog|FormFill|{requestorId}|{sha256(url)}
 * Cache value: { formHash, fillActions } as JSON
 *
 * Degrades gracefully when Redis is unavailable — all operations become no-ops and the
 * step falls through to the AI call.
 */
export class AiFormCache {
  private getAsync: ((key: string) => Promise<string | null>) | null = null;
  private setAsync: ((key: string, ttl: number, value: string) => Promise<any>) | null = null;

  // 8 days in seconds — covers weekly-cadence tests with a full day of margin
  private static readonly TTL_SECONDS = 691200;

  constructor(private redisClient: any) {
    if (redisClient) {
      try {
        this.getAsync = promisify(redisClient.get).bind(redisClient);
        this.setAsync = promisify(redisClient.setex).bind(redisClient);
      } catch (err) {
        console.log('[AiFormCache] Failed to promisify Redis methods:', err);
      }
    }
  }

  private buildKey(requestorId: string, url: string): string {
    const urlHash = crypto.createHash('sha256').update(url).digest('hex');
    return `WebCog|FormFill|${requestorId}|${urlHash}`;
  }

  public async get(requestorId: string, url: string): Promise<CachedFormFill | null> {
    if (!this.getAsync) return null;
    try {
      const stored = await this.getAsync(this.buildKey(requestorId, url));
      if (stored) {
        return JSON.parse(stored) as CachedFormFill;
      }
      return null;
    } catch (err) {
      console.log('[AiFormCache] Redis get error (non-fatal):', err);
      return null;
    }
  }

  public async set(requestorId: string, url: string, formHash: string, fillActions: AiCachedFillAction[]): Promise<void> {
    if (!this.setAsync) return;
    try {
      const value: CachedFormFill = { formHash, fillActions };
      await this.setAsync(this.buildKey(requestorId, url), AiFormCache.TTL_SECONDS, JSON.stringify(value));
    } catch (err) {
      console.log('[AiFormCache] Redis set error (non-fatal):', err);
    }
  }
}
