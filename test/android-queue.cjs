const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const html=fs.readFileSync('index.html','utf8');
const code=html.slice(html.indexOf('let googleDrainBusy=false;'),html.indexOf("window.addEventListener('online',drainGoogleQueue);"));
assert.ok(code);
const shared={url:'',key:'',q:[{eventId:'id',revisionId:'revision',schemaVersion:2}],configured:[],enqueued:[],fail:false};
const context={window:{Capacitor:{Plugins:{GoogleSync:{configure:async o=>(shared.configured.push(o),{ok:true}),enqueue:async o=>(shared.enqueued.push(o),{ok:!shared.fail})}}}},GOOGLE_URL_STORE:'url',GOOGLE_KEY_STORE:'key',googleRead:k=>shared[k]||'',googleWrite:()=>true,googleQueue:()=>structuredClone(shared.q),saveGoogleQueue:q=>(shared.q=q,true),AbortController,setTimeout,clearTimeout};
vm.createContext(context);vm.runInContext(code,context);
(async()=>{await context.drainGoogleQueue();assert.equal(shared.configured.length,0);assert.equal(shared.q.length,1);shared.url='https://example.test';shared.key='demo';shared.fail=true;await context.drainGoogleQueue();assert.equal(shared.q.length,1);shared.fail=false;await context.drainGoogleQueue();assert.equal(shared.q.length,0);assert.equal(shared.enqueued.length,2);console.log('Stored configuration and pending jobs are preserved');})().catch(e=>{console.error(e);process.exitCode=1});
