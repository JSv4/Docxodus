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

for(const worker of [false,true]) {
  for(const sample of cases) test(`${sample.name} ownership (${worker?'worker':'main thread'})`,async({page})=>{
    const source=revisionIdentityDocx(sample.author,sample.account);
    await page.goto(worker?'/worker-test-harness.html':'/test-harness.html');
    if(!worker) await page.waitForFunction(()=>(window as any).DocxodusReady===true);
    const result=await page.evaluate(async({bytes,worker,sample,account,proposal})=>{
      if(worker) await (window as any).createDocxodusWorker();
      const api=worker?(window as any).DocxodusWorker:(window as any).Docxodus;
      const settings={trackedChanges:'render_inline',revisionAuthor:'Pat',
        revisionIdentities:sample.identities?[{providerId:'AD',userId:account}]:null};
      const session=worker?await api.openDocxSession(new Uint8Array(bytes),settings)
        :api.openTypedSession(new Uint8Array(bytes),settings);
      try {
        const name=await session.getRevisionAuthorNameStatus(sample.author);
        const target=await session.findByText(proposal);
        const edit=await session.deleteBlock(target.id);
        return {name,edit,revisions:await session.listRevisions(),saved:Array.from(await session.save()) as number[]};
      } finally {await session.close();if(worker) api.terminate();}
    },{bytes:Array.from(source),worker,sample,account:OUR_ACCOUNT,proposal:IDENTITY_PROPOSAL});
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

  test(`mid-session configuration and an undoable identity claim (${worker?'worker':'main thread'})`,async({page})=>{
    await page.goto(worker?'/worker-test-harness.html':'/test-harness.html');
    if(!worker) await page.waitForFunction(()=>(window as any).DocxodusReady===true);
    const source=revisionIdentityDocx('Earlier',null);
    const result=await page.evaluate(async({bytes,worker,account})=>{
      if(worker) await (window as any).createDocxodusWorker();
      const api=worker?(window as any).DocxodusWorker:(window as any).Docxodus;
      const options={trackedChanges:'render_inline',revisionAuthor:'Pat'};
      const session=worker?await api.openDocxSession(new Uint8Array(bytes),options):api.openTypedSession(new Uint8Array(bytes),options);
      try {
        await session.setRevisionIdentities([{providerId:'AD',userId:account}]);
        const before=await session.getRevisionAuthorNameStatus();
        const anchor=await session.findByText('Before.');
        const edit=await session.replaceText(anchor.id,'Edited.');
        const after=await session.getRevisionAuthorNameStatus();
        const saved=Array.from(await session.save()) as number[];
        const undone=await session.undo(),undoStatus=await session.getRevisionAuthorNameStatus();
        const redone=await session.redo(),redoStatus=await session.getRevisionAuthorNameStatus();
        await session.setRevisionIdentities(null);
        return {before,after,edit,saved,undone,undoStatus,redone,redoStatus};
      } finally {await session.close();if(worker) api.terminate();}
    },{bytes:Array.from(source),worker,account:OUR_ACCOUNT});
    expect(result.edit.success).toBe(true);
    expect(result.before.status).toBe('free');expect(result.after.status).toBe('own');
    expect(result.undone).toBe(true);expect(result.undoStatus.status).toBe('free');
    expect(result.redone).toBe(true);expect(result.redoStatus.status).toBe('own');
    await test.info().attach('result.docx',{body:Buffer.from(result.saved),contentType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  });
}
