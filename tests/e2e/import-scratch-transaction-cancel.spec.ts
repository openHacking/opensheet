import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const format of ['validation', 'json', 'binary'] as const) {
  for (const phase of ['before', 'after'] as const) {
    test(`cancelling a ${format} scratch write ${phase} commit preserves the old workbook`, async ({
      page,
    }) => {
      const modulePath = (name: string) =>
        `/@fs${fileURLToPath(new URL(`../../packages/core/src/${name}.ts`, import.meta.url))}`;
      await page.route('**/import-scratch-transaction-cancel.worker.js', (route) =>
        route.fulfill({
          contentType: 'text/javascript',
          body: `
      import { Engine } from ${JSON.stringify(modulePath('engine'))};
      import { Database } from ${JSON.stringify(modulePath('storage'))};
      import { createSnapshot } from ${JSON.stringify(modulePath('model'))};
      import { encodeWorkbookFile } from ${JSON.stringify(modulePath('file-codec'))};
      import { encodeBlock } from ${JSON.stringify(modulePath('block-codec'))};
      import { writeBinary } from ${JSON.stringify(modulePath('binary-file'))};
      import { upgradeValidationFile } from ${JSON.stringify(modulePath('validation-file-migration'))};
      (async () => {
        const traceId = crypto.randomUUID().slice(0,8);
        const progress = phase => self.postMessage({progress:${JSON.stringify(format)} + '-' + ${JSON.stringify(phase)} + ':' + traceId + ':' + phase});
        progress('fixture');
        const snapshot = createSnapshot({sheets:[{name:'Data',rows:130,columns:1}]});
        snapshot.sheets[0].cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:5}};
        snapshot.sheets[0].cells['64:0'] = {rowId:64,columnId:0,input:{type:'number',value:6}};
        snapshot.sheets[0].cells['128:0'] = {rowId:128,columnId:0,input:{type:'number',value:7}};
        const format = ${JSON.stringify(format)};
        const phase = ${JSON.stringify(phase)};
        if (${JSON.stringify(!!process.env.OPENSHEET_TRACE_IMPORT_CANCEL)} && format === 'binary') {
          const NativeDecompressionStream = globalThis.DecompressionStream;
          let decoderId = 0;
          globalThis.DecompressionStream = class extends NativeDecompressionStream {
            constructor(...args) {
              super(...args);
              const id = ++decoderId, readable = this.readable;
              const getReader = readable.getReader.bind(readable);
              readable.getReader = (...args) => {
                const reader = getReader(...args), read = reader.read.bind(reader), cancel = reader.cancel.bind(reader), release = reader.releaseLock.bind(reader);
                let n = 0;
                reader.read = async (...args) => {
                  const index = ++n;
                  progress('decoder-'+id+'-read-'+index+'-start');
                  const result = await read(...args);
                  progress('decoder-'+id+'-read-'+index+'-'+(result.done?'done':'bytes-'+result.value.length));
                  return result;
                };
                reader.cancel = async (...args) => {
                  progress('decoder-'+id+'-cancel-start');
                  const result = await cancel(...args);
                  progress('decoder-'+id+'-cancel-done');
                  return result;
                };
                reader.releaseLock = () => {
                  progress('decoder-'+id+'-release-start');
                  const result = release();
                  progress('decoder-'+id+'-release-done');
                  return result;
                };
                return reader;
              };
            }
          };
          const originalStream = Blob.prototype.stream;
          let streamId = 0;
          Blob.prototype.stream = function() {
            const id = ++streamId, stream = originalStream.call(this);
            progress('blob-stream-'+id+'-size-'+this.size);
            const getReader = stream.getReader.bind(stream);
            stream.getReader = (...args) => {
              const reader = getReader(...args), read = reader.read.bind(reader);
              let n = 0;
              reader.read = async (...args) => {
                const index = ++n;
                progress('blob-'+id+'-read-'+index+'-start');
                const result = await read(...args);
                progress('blob-'+id+'-read-'+index+'-'+(result.done?'done':'bytes-'+result.value.length));
                return result;
              };
              return reader;
            };
            return stream;
          };
        }
        const source = encodeWorkbookFile(snapshot);
        let backup;
        if (format === 'binary') {
          const metadata = {...source,sheets:source.sheets.map(sheet => ({...sheet,blocks:[]}))};
          const parts = [];
          for await (const part of writeBinary(metadata, (async function* () {
            for (const row of [0,64,128]) {
              const cell = snapshot.sheets[0].cells[row+':0'];
              yield {sheet:0,row:row/64,column:0,bytes:encodeBlock({[row+':0']:cell}, [], [], {})};
            }
          })())) parts.push(part);
          backup = new Blob(parts);
        } else backup = new Blob([JSON.stringify(format === 'validation' ? upgradeValidationFile(source) : source)]);
        const runImport = (engine) => format === 'validation' ? engine.importRuleFreeValidationBackup(backup)
          : format === 'json' ? engine.importJSON(backup) : engine.importBinary(backup);
        const database = 'scratch-cancel-' + crypto.randomUUID();
        const db = await Database.open(database);
        const engine = new Engine(db,snapshot.workbookId);
        progress('open');
        await engine.open(snapshot);
        try {
          const sheet = snapshot.sheets[0].id;
          await engine.execute([{type:'core.cells.set',payload:{sheetId:sheet,cells:[{row:0,column:0,input:{type:'number',value:9}}]}}]);
          const before = engine.view();
          const capture = async () => {
            const tx = db.db.transaction(['heads','records','sizes','scratch']);
            const get = request => new Promise((resolve,reject) => { request.onsuccess=()=>resolve(request.result); request.onerror=()=>reject(request.error); });
            const range = IDBKeyRange.bound(snapshot.workbookId + '/', snapshot.workbookId + '/\\uffff');
            const requests = [tx.objectStore('heads').get(snapshot.workbookId), ...['records','sizes','scratch'].flatMap(name => [tx.objectStore(name).getAllKeys(range),tx.objectStore(name).getAll(range)])];
            return JSON.stringify(await Promise.all(requests.map(get)), (_key,value) => value instanceof Uint8Array ? Array.from(value) : value);
          };
          const rawBefore = await capture();
          progress('first-import');
          const originalPut = IDBObjectStore.prototype.put;
          let reached = false, transactionAborted = false, transactionCommitted = false, error = '', attempts = 0, priorCommits = 0;
          IDBObjectStore.prototype.put = function(...args) {
            const request = originalPut.apply(this,args);
            if (this.name === 'scratch' && !reached) {
              const transaction = this.transaction;
              progress('scratch-put-'+(attempts+1));
              if (++attempts < 3) transaction.addEventListener('complete', () => { priorCommits++; progress('scratch-committed-'+priorCommits); }, {once:true});
              else {
                transaction.addEventListener('abort', () => { transactionAborted = true; progress('scratch-aborted'); }, {once:true});
                transaction.addEventListener('complete', () => { transactionCommitted = true; progress('scratch-committed-3'); }, {once:true});
                if (phase === 'before') request.addEventListener('success', () => { reached = true; progress('cancel-requested'); engine.cancel(); }, {once:true});
                else transaction.addEventListener('complete', () => { reached = true; progress('cancel-requested'); engine.cancel(); }, {once:true});
              }
            }
            return request;
          };
          try { await runImport(engine); }
          catch (reason) { error = reason.code; }
          finally { IDBObjectStore.prototype.put = originalPut; }
          progress('verify-rollback');
          engine.begin();
          const after = engine.view();
          const value = (await engine.cell(sheet,0,0)).input.value;
          const scratch = await db.hasAny('scratch', IDBKeyRange.bound(snapshot.workbookId + '/', snapshot.workbookId + '/\\uffff'));
          const rawUnchanged = (await capture()) === rawBefore;
          progress('history');
          await engine.history(false);
          const undone = (await engine.cell(sheet,0,0)).input.value;
          await engine.history(true);
          const redone = (await engine.cell(sheet,0,0)).input.value;
          engine.close();
          progress('reopen');
          const reopenedDB = await Database.open(database);
          const reopenedEngine = new Engine(reopenedDB,snapshot.workbookId);
          let reopened, retried;
          try {
            await reopenedEngine.open();
            reopened = (await reopenedEngine.cell(sheet,0,0)).input.value;
            progress('retry-import');
            await runImport(reopenedEngine);
            progress('retry-read');
            retried = await Promise.all([0,64,128].map(async row => (await reopenedEngine.cell(sheet,row,0)).input.value));
          } finally { reopenedEngine.close(); }
          self.postMessage({reached,attempts,priorCommits,transactionAborted,transactionCommitted,error,beforeRevision:before.revision,afterRevision:after.revision,beforeBytes:before.bytes,afterBytes:after.bytes,value,scratch,rawUnchanged,undone,redone,reopened,retried});
        } finally { engine.close(); }
      })().catch(error => self.postMessage({error:String(error),stack:error.stack}));
    `,
        }),
      );
      await page.goto('/');
      page.on('console', (message) => {
        if (
          process.env.OPENSHEET_TRACE_IMPORT_CANCEL &&
          message.text().startsWith('scratch-phase:')
        )
          console.log(message.text());
      });
      const result = await page.evaluate(
        () =>
          new Promise<any>((resolve, reject) => {
            const worker = new Worker('/import-scratch-transaction-cancel.worker.js', {
              type: 'module',
            });
            let phase = 'worker-start';
            const timer = setTimeout(() => {
              worker.terminate();
              reject(new Error(`Worker timeout during ${phase}`));
            }, 10000);
            worker.onerror = (error) => {
              clearTimeout(timer);
              worker.terminate();
              reject(new Error(error.message));
            };
            worker.onmessage = ({ data }) => {
              if (data.progress) {
                phase = data.progress;
                console.log(`scratch-phase:${phase}`);
                return;
              }
              clearTimeout(timer);
              worker.terminate();
              resolve(data);
            };
          }),
      );
      expect(result).toMatchObject({
        reached: true,
        attempts: 3,
        priorCommits: 2,
        transactionAborted: phase === 'before',
        transactionCommitted: phase === 'after',
        error: 'ABORTED',
        value: 9,
        scratch: false,
        rawUnchanged: true,
        undone: 5,
        redone: 9,
        reopened: 9,
        retried: [5, 6, 7],
      });
      expect(result.afterRevision).toBe(result.beforeRevision);
      expect(result.afterBytes).toBe(result.beforeBytes);
    });
  }
}
