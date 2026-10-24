import {expect,test} from '@playwright/test';
import type {DocxodusWasmExports,RevisionListEntry} from '../src/types.js';
import {openDocxSession} from '../src/session.js';
import {IDENTITY_PROPOSAL, OTHER_ACCOUNT, OUR_ACCOUNT, revisionIdentityDocx} from './docx-revision-identity-fixture.js';

const cases = [
  {name:'namesake account',author:'Pat',account:OTHER_ACCOUNT,identities:true,own:false},
  {name:'unattributed namesake',author:'Pat',account:null,identities:true,own:false},
  {name:'UPN alias',author:'Word Pat',account:'S::PAT@OURFIRM.EXAMPLE::33333333-3333-3333-3333-333333333333',identities:true,own:true},
  {name:'object-id alias',author:'Word Pat',account:'S::former@ourfirm.example::22222222-2222-2222-2222-222222222222',identities:true,own:true},
  {name:'registered same name',author:'Pat',account:OUR_ACCOUNT,identities:true,own:true},
  {name:'legacy name control',author:'Pat',account:OTHER_ACCOUNT,identities:false,own:true},
];

test('account settings refuse an older bridge before opening a session',()=>{
  let opened=0;
  const old={DocxSessionBridge:{OpenSession:()=>{opened++;return 1;},CloseSession:()=>{}}} as unknown as DocxodusWasmExports;
  expect(()=>openDocxSession(revisionIdentityDocx(),old,{
    revisionIdentities:[{providerId:'AD',userId:OUR_ACCOUNT}],
  })).toThrow('Account-based revision ownership');
  expect(opened).toBe(0);
});

test('a bridge without account support keeps the legacy defaults and empty opt-outs',()=>{
  let opened=0;
  const old={DocxSessionBridge:{OpenSession:()=>++opened,CloseSession:()=>{}}} as unknown as DocxodusWasmExports;
  for(const revisionIdentities of [undefined,null,[]]) {
    const session=openDocxSession(revisionIdentityDocx(),old,{revisionIdentities});
    session.setRevisionIdentities(null);
    session.close();
  }
  expect(opened).toBe(3);
});

test('the editor forwards account settings and its configuration/query methods',async({page})=>{
  await page.goto('/test-harness.html');
  await page.waitForFunction(()=>(window as any).DocxodusReady===true);
  const result=await page.evaluate(({bytes,account})=>{
    const api=(window as any).Docxodus;
    const container=document.createElement('div');document.body.append(container);
    const editor=api.DocxEditor.open(container,new Uint8Array(bytes),api,{
      trackedChanges:1,revisionAuthor:'Pat',comments:false,
      revisionIdentities:[{providerId:'AD',userId:account}],
    });
    try {
      const initial=editor.getRevisionAuthorNameStatus();
      editor.setRevisionIdentities(null);
      const legacy=editor.getRevisionAuthorNameStatus();
      editor.setRevisionIdentities([{providerId:'AD',userId:account}]);
      const restored=editor.getRevisionAuthorNameStatus();
      editor.setRevisionAuthor('Pat (our firm)');
      const unique=editor.getRevisionAuthorNameStatus();
      return {initial,legacy,restored,unique};
    } finally {editor.close();container.remove();}
  },{bytes:Array.from(revisionIdentityDocx()),account:OUR_ACCOUNT});
  expect(result.initial.status).toBe('taken');expect(result.legacy.status).toBe('own');
  expect(result.restored.status).toBe('taken');expect(result.unique.status).toBe('free');
});

for(const sample of cases) test(`${sample.name} deletion ownership (main thread)`,async({page})=>{
  const source=revisionIdentityDocx(sample.author,sample.account);
  await page.goto('/test-harness.html');
  await page.waitForFunction(()=>(window as any).DocxodusReady===true);
  const result=await page.evaluate(({bytes,sample,account,proposal})=>{
    const api=(window as any).Docxodus;
    const settings={trackedChanges:'render_inline',revisionAuthor:'Pat',
      revisionIdentities:sample.identities?[{providerId:'AD',userId:account}]:null};
    const session=(window as any).DocxodusSession.openDocxSession(new Uint8Array(bytes),api,settings);
    try {
      const name=session.getRevisionAuthorNameStatus(sample.author);
      const target=session.findByText(proposal);
      const edit=session.deleteBlock(target.id);
      return {name,edit,revisions:session.listRevisions(),saved:Array.from(session.save()) as number[]};
    } finally {session.close();}
  },{bytes:Array.from(source),sample,account:OUR_ACCOUNT,proposal:IDENTITY_PROPOSAL});
  expect(result.name).toEqual({author:sample.author,status:sample.own?'own':'taken'});
  expect(result.edit.success).toBe(true);
  if(sample.own) expect(result.revisions).toEqual([]);
  else {
    expect((result.revisions as RevisionListEntry[]).some(r=>r.type==='insert')).toBe(true);
    expect((result.revisions as RevisionListEntry[]).some(r=>r.type==='delete')).toBe(true);
  }
  await test.info().attach('source.docx',{body:Buffer.from(source),contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  await test.info().attach('result.docx',{body:Buffer.from(result.saved),contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
});

test('mid-session configuration and an undoable identity claim (main thread)',async({page})=>{
  await page.goto('/test-harness.html');
  await page.waitForFunction(()=>(window as any).DocxodusReady===true);
  const result=await page.evaluate(({bytes,account})=>{
    const api=(window as any).Docxodus;
    const session=(window as any).DocxodusSession.openDocxSession(new Uint8Array(bytes),api,
      {trackedChanges:'render_inline',revisionAuthor:'Pat'});
    try {
      session.setRevisionIdentities([{providerId:'AD',userId:account}]);
      const before=session.getRevisionAuthorNameStatus();
      const anchor=session.findByText('Before.');
      const edit=session.replaceText(anchor.id,'Edited.');
      const after=session.getRevisionAuthorNameStatus();
      const saved=Array.from(session.save()) as number[];
      const undone=session.undo(),undoStatus=session.getRevisionAuthorNameStatus();
      const redone=session.redo(),redoStatus=session.getRevisionAuthorNameStatus();
      session.setRevisionIdentities(null);
      return {before,after,edit,saved,undone,undoStatus,redone,redoStatus};
    } finally {session.close();}
  },{bytes:Array.from(revisionIdentityDocx('Earlier',null)),account:OUR_ACCOUNT});
  expect(result.edit.success).toBe(true);
  expect(result.before.status).toBe('free');expect(result.after.status).toBe('own');
  expect(result.undone).toBe(true);expect(result.undoStatus.status).toBe('free');
  expect(result.redone).toBe(true);expect(result.redoStatus.status).toBe('own');
  await test.info().attach('result.docx',{body:Buffer.from(result.saved),contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
});

for(const sample of cases) test(`${sample.name} account setting and name query (worker)`,async({page})=>{
  await page.goto('/worker-test-harness.html');
  const result=await page.evaluate(async({bytes,sample,account})=>{
    await (window as any).createDocxodusWorker();
    const api=(window as any).DocxodusWorker as import('../src/worker-proxy.js').WorkerDocxodus;
    const session=await api.openDocxSession(new Uint8Array(bytes),{
      revisionAuthor:'Pat',revisionIdentities:sample.identities?[{providerId:'AD',userId:account}]:null,
    });
    try {
      return {name:await session.getRevisionAuthorNameStatus(sample.author),
        mainThreadRuntime:typeof (window as any).Docxodus!=='undefined'};
    } finally {await session.close();api.terminate();}
  },{bytes:Array.from(revisionIdentityDocx(sample.author,sample.account)),sample,account:OUR_ACCOUNT});
  expect(result.name).toEqual({author:sample.author,status:sample.own?'own':'taken'});
  expect(result.mainThreadRuntime).toBe(false);
});

test('worker configuration resets and invalid identities preserve the previous configuration',async({page})=>{
  await page.goto('/worker-test-harness.html');
  const result=await page.evaluate(async({bytes,account})=>{
    await (window as any).createDocxodusWorker();
    const api=(window as any).DocxodusWorker as import('../src/worker-proxy.js').WorkerDocxodus;
    const session=await api.openDocxSession(new Uint8Array(bytes),{revisionAuthor:'Pat'});
    try {
      const legacy=await session.getRevisionAuthorNameStatus();
      await session.setRevisionIdentities([{providerId:'AD',userId:account}]);
      const configured=await session.getRevisionAuthorNameStatus();
      let invalidRejected=false;
      try {await session.setRevisionIdentities([{providerId:'',userId:account}]);}
      catch {invalidRejected=true;}
      const afterInvalid=await session.getRevisionAuthorNameStatus();
      const freeName=await session.getRevisionAuthorNameStatus('Pat (our firm)');
      await session.setRevisionIdentities([]);
      const empty=await session.getRevisionAuthorNameStatus();
      await session.setRevisionIdentities([{providerId:'AD',userId:account}]);
      await session.setRevisionIdentities(null);
      const cleared=await session.getRevisionAuthorNameStatus();
      return {legacy,configured,invalidRejected,afterInvalid,freeName,empty,cleared};
    } finally {await session.close();api.terminate();}
  },{bytes:Array.from(revisionIdentityDocx()),account:OUR_ACCOUNT});
  expect(result.legacy.status).toBe('own');expect(result.configured.status).toBe('taken');
  expect(result.invalidRejected).toBe(true);expect(result.afterInvalid.status).toBe('taken');
  expect(result.freeName.status).toBe('free');expect(result.empty.status).toBe('own');
  expect(result.cleared.status).toBe('own');
});

test('worker identity configuration stays with its session handle',async({page})=>{
  await page.goto('/worker-test-harness.html');
  const result=await page.evaluate(async({bytes,account})=>{
    await (window as any).createDocxodusWorker();
    const api=(window as any).DocxodusWorker as import('../src/worker-proxy.js').WorkerDocxodus;
    const legacy=await api.openDocxSession(new Uint8Array(bytes),{revisionAuthor:'Pat'});
    const configured=await api.openDocxSession(new Uint8Array(bytes),{revisionAuthor:'Pat'});
    try {
      await configured.setRevisionIdentities([{providerId:'AD',userId:account}]);
      const statuses=await Promise.all([legacy.getRevisionAuthorNameStatus(),configured.getRevisionAuthorNameStatus()]);
      await configured.close();
      let closedRejected=false;
      try {await configured.getRevisionAuthorNameStatus();} catch {closedRejected=true;}
      return {statuses,closedRejected};
    } finally {await legacy.close();api.terminate();}
  },{bytes:Array.from(revisionIdentityDocx()),account:OUR_ACCOUNT});
  expect(result.statuses.map(info=>info.status)).toEqual(['own','taken']);
  expect(result.closedRejected).toBe(true);
});
