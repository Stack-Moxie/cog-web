import * as crypto from 'crypto';
import * as chai from 'chai';
import { default as sinon } from 'ts-sinon';
import * as sinonChai from 'sinon-chai';
import 'mocha';

import { AiFormCache, AiCachedFillAction } from '../../src/client/ai-form-cache';

chai.use(sinonChai);

describe('AiFormCache', () => {
  const expect = chai.expect;

  const REQUESTOR_ID = 'org-123';
  const URL = 'https://example.com/form';
  const FORM_HASH = 'abc123hash';
  const FILL_ACTIONS: AiCachedFillAction[] = [
    { selector: '#email', value: 'test@example.com', inputType: 'text' },
    { selector: 'button[type="submit"]', value: 'submit', inputType: 'click' },
  ];

  function expectedKey(): string {
    const urlHash = crypto.createHash('sha256').update(URL).digest('hex');
    return `WebCog|FormFill|${REQUESTOR_ID}|${urlHash}`;
  }

  // ── No Redis client ───────────────────────────────────────────────────────────

  describe('when no Redis client is provided', () => {
    let cache: AiFormCache;

    beforeEach(() => {
      cache = new AiFormCache(null);
    });

    it('get() returns null without throwing', async () => {
      const result = await cache.get(REQUESTOR_ID, URL);
      expect(result).to.be.null;
    });

    it('set() is a no-op without throwing', async () => {
      await expect(cache.set(REQUESTOR_ID, URL, FORM_HASH, FILL_ACTIONS)).to.not.be.rejected;
    });
  });

  // ── With a Redis client ───────────────────────────────────────────────────────

  describe('when a Redis client is provided', () => {
    let cache: AiFormCache;
    let redisGetStub: sinon.SinonStub;
    let redisSetexStub: sinon.SinonStub;
    let redisClientStub: any;

    beforeEach(() => {
      // Redis v3 uses callbacks; promisify wraps them.
      // Stubs here use the callback style (last arg is cb).
      redisGetStub = sinon.stub().callsFake((_key: string, cb: Function) => cb(null, null));
      redisSetexStub = sinon.stub().callsFake((_key: string, _ttl: number, _val: string, cb: Function) => cb(null, 'OK'));
      redisClientStub = { get: redisGetStub, setex: redisSetexStub };
      cache = new AiFormCache(redisClientStub);
    });

    // ── get() ─────────────────────────────────────────────────────────────────

    it('get() returns null on a cache miss', async () => {
      const result = await cache.get(REQUESTOR_ID, URL);
      expect(result).to.be.null;
      expect(redisGetStub).to.have.been.calledOnce;
    });

    it('get() returns parsed CachedFormFill on a cache hit', async () => {
      const stored = JSON.stringify({ formHash: FORM_HASH, fillActions: FILL_ACTIONS });
      redisGetStub.callsFake((_key: string, cb: Function) => cb(null, stored));

      const result = await cache.get(REQUESTOR_ID, URL);

      expect(result).to.not.be.null;
      expect(result!.formHash).to.equal(FORM_HASH);
      expect(result!.fillActions).to.deep.equal(FILL_ACTIONS);
    });

    it('get() returns null (non-fatal) when Redis throws', async () => {
      redisGetStub.callsFake((_key: string, cb: Function) => cb(new Error('Redis connection refused')));

      const result = await cache.get(REQUESTOR_ID, URL);
      expect(result).to.be.null;
    });

    // ── set() ─────────────────────────────────────────────────────────────────

    it('set() calls setex with the correct key, TTL, and serialised value', async () => {
      await cache.set(REQUESTOR_ID, URL, FORM_HASH, FILL_ACTIONS);

      expect(redisSetexStub).to.have.been.calledOnce;
      const [key, ttl, value] = redisSetexStub.firstCall.args;

      expect(key).to.equal(expectedKey());
      expect(ttl).to.equal(691200); // 8 days
      const parsed = JSON.parse(value);
      expect(parsed.formHash).to.equal(FORM_HASH);
      expect(parsed.fillActions).to.deep.equal(FILL_ACTIONS);
    });

    it('set() is non-fatal when Redis throws', async () => {
      redisSetexStub.callsFake((_k: string, _t: number, _v: string, cb: Function) =>
        cb(new Error('ECONNREFUSED')),
      );

      await expect(cache.set(REQUESTOR_ID, URL, FORM_HASH, FILL_ACTIONS)).to.not.be.rejected;
    });

    // ── Key stability ─────────────────────────────────────────────────────────

    it('get() uses the same cache key as set()', async () => {
      await cache.set(REQUESTOR_ID, URL, FORM_HASH, FILL_ACTIONS);
      await cache.get(REQUESTOR_ID, URL);

      const setKey = redisSetexStub.firstCall.args[0];
      const getKey = redisGetStub.firstCall.args[0];
      expect(getKey).to.equal(setKey);
    });

    it('produces a different key for a different URL', async () => {
      await cache.get(REQUESTOR_ID, URL);
      await cache.get(REQUESTOR_ID, 'https://other.com/form');

      const key1 = redisGetStub.firstCall.args[0];
      const key2 = redisGetStub.secondCall.args[0];
      expect(key1).to.not.equal(key2);
    });

    it('produces a different key for a different requestorId', async () => {
      await cache.get('org-aaa', URL);
      await cache.get('org-bbb', URL);

      const key1 = redisGetStub.firstCall.args[0];
      const key2 = redisGetStub.secondCall.args[0];
      expect(key1).to.not.equal(key2);
    });

    it('cache key contains the sha256 hash of the URL, not the raw URL', async () => {
      await cache.get(REQUESTOR_ID, URL);
      const key = redisGetStub.firstCall.args[0];
      expect(key).to.not.include(URL);
      const expectedUrlHash = crypto.createHash('sha256').update(URL).digest('hex');
      expect(key).to.include(expectedUrlHash);
    });

    it('cache key includes the requestorId verbatim', async () => {
      await cache.get(REQUESTOR_ID, URL);
      const key = redisGetStub.firstCall.args[0];
      expect(key).to.include(REQUESTOR_ID);
    });

    it('cache key starts with the WebCog|FormFill namespace', async () => {
      await cache.get(REQUESTOR_ID, URL);
      const key = redisGetStub.firstCall.args[0];
      expect(key).to.match(/^WebCog\|FormFill\|/);
    });
  });
});
