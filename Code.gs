/** 岡谷市 電力供給計画／Google Apps Script (V8)
 * スプレッドシートに紐づけたGASとして利用してください。
 * 管理者が最初に setupApp() を実行します。
 */
const SHEET = Object.freeze({CONFIG:'設定', SOURCES:'発電資料', RECORDS:'学習記録'});
const FIELDS = ['timestamp','classCode','studentCode','stage','planJson','reason','comparison','constraints','disaster','revision','dataYear','totalPercent','initialPlanJson','changeReason','disasterCountermeasure','outageImpact'];
const INITIAL_CONFIG = [
 ['key','value','explanation'],
 ['demandMWh',307794,'岡谷市 年間電気使用量の教材用参照値（MWh/年）。2022年度推計値を後年比較に代用。'],
 ['demandLabel','2022年度推計（2023年度比較に代用）','実測値ではありません。'],
 ['demandSource','https://www.env.go.jp/policy/local_keikaku/data/karte/pdf/20204.pdf','環境省 自治体排出量カルテ'],
 ['potentialLabel','岡谷市 第5次環境基本計画 表4-3','導入済み量ではなく、年間発電量ポテンシャル。'],
 ['potentialSource','https://www.city.okaya.lg.jp/material/files/group/48/5thPlanhonpen.pdf','岡谷市 第5次環境基本計画'],
 ['classPassphrase','','任意：空欄ならクラスコードのみ。秘密情報に使わない。'],
 ['allowSubmissions','TRUE','FALSEにすると記録送信を受け付けません。']
];
const INITIAL_SOURCES = [
 ['id','name','energy','merits','limits','potentialMWh','color','guide','sourceUrl'],
 ['solar','太陽光','光 → 電気','屋根などにも設置でき、運転時の直接CO₂排出が少ない','夜間は発電できず天候に左右される',407039,'#e5a634','日射、建物の強度、費用、自立運転・蓄電池の条件','https://www.city.okaya.lg.jp/material/files/group/48/5thPlanhonpen.pdf'],
 ['wind','風力','風の運動 → 電気','運転時の直接CO₂排出が少ない','風況に左右され、騒音・生態系・設備管理への配慮が必要',29061,'#4b9f99','風況、自然環境、送電設備との接続','https://www.city.okaya.lg.jp/material/files/group/48/5thPlanhonpen.pdf'],
 ['hydro','水力','水の位置・運動 → 電気','方式によって比較的安定して発電できる','水量と落差、河川利用などに制約がある',2719,'#4e8bc9','流量、落差、河川管理、費用対効果','https://www.city.okaya.lg.jp/material/files/group/48/5thPlanhonpen.pdf'],
 ['thermal','火力','燃料の化学 → 熱 → 運動 → 電気','出力調整がしやすい方式がある','燃料調達・CO₂排出・立地などが課題','', '#c87863','用地、燃料供給、排出・排熱対策、災害時の燃料',''],
 ['nuclear','原子力','核分裂の熱 → 運動 → 電気','運転時の直接CO₂排出が少ない','厳格な安全対策・法規制・放射性廃棄物対応が必要','','#877eb0','立地規制、冷却水、安全・防災、廃棄物への対応',''],
 ['geothermal','地熱','地下の熱 → 運動 → 電気','発電に適した資源があれば天候に左右されにくい','高温の地下資源・掘削費・環境影響の調査が必要','','#a87b60','高温資源の調査（地中熱利用とは別）、掘削、温泉への影響',''],
 ['external','市外から供給','市外の発電所 → 送電・変電・配電 → 市内','岡谷市以外の電源も利用できる','送配電網の被災時に供給が滞る場合がある','','#6e819b','発電方式そのものではない。市外の電源構成、送電・配電を調べる','']
];
function setupApp() {
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 if(!ss) throw new Error('スプレッドシートに紐づくGASプロジェクトで実行してください。');
 createOnce_(ss,SHEET.CONFIG,INITIAL_CONFIG);
 createOnce_(ss,SHEET.SOURCES,INITIAL_SOURCES);
 createOnce_(ss,SHEET.RECORDS,[FIELDS]);
 upgradeRecords_(ss.getSheetByName(SHEET.RECORDS));
 const r=ss.getSheetByName(SHEET.RECORDS);
 r.setFrozenRows(1);
 r.getRange(1,1,1,FIELDS.length).setBackground('#174c62').setFontColor('white').setFontWeight('bold');
 r.autoResizeColumns(1,Math.min(FIELDS.length,4));
 r.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
 return '設定・発電資料・学習記録の3シートを確認してください。';
}
function upgradeRecords_(sheet){
 const present=sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0].map(String);
 const missing=FIELDS.filter(f=>!present.includes(f));
 if(missing.length){let end=sheet.getLastColumn();if(end+missing.length>sheet.getMaxColumns())sheet.insertColumnsAfter(sheet.getMaxColumns(),end+missing.length-sheet.getMaxColumns());sheet.getRange(1,end+1,1,missing.length).setValues([missing]).setBackground('#174c62').setFontColor('white').setFontWeight('bold');}
}
function createOnce_(ss,name,rows){
 let sh=ss.getSheetByName(name); if(sh) return;
 sh=ss.insertSheet(name); sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
 sh.setFrozenRows(1);sh.getRange(1,1,1,rows[0].length).setBackground('#174c62').setFontColor('white').setFontWeight('bold');
 sh.autoResizeColumns(1,rows[0].length);
}
function doGet(){return HtmlService.createHtmlOutputFromFile('Index').setTitle('岡谷市 電力供給計画').addMetaTag('viewport','width=device-width,initial-scale=1');}
function getAppData(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();
 const cfg=Object.fromEntries(ss.getSheetByName(SHEET.CONFIG).getDataRange().getValues().slice(1).map(r=>[String(r[0]),r[1]]));
 const sh=ss.getSheetByName(SHEET.SOURCES), rows=sh.getDataRange().getValues();
 const headers=rows.shift();
 const sources=rows.filter(r=>r[0]).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]]))).map(s=>({ ...s,potentialMWh:s.potentialMWh===''?null:Number(s.potentialMWh),sourceUrl:String(s.sourceUrl||'')}));
 return {demandMWh:Number(cfg.demandMWh),demandLabel:String(cfg.demandLabel),demandSource:String(cfg.demandSource),potentialLabel:String(cfg.potentialLabel),potentialSource:String(cfg.potentialSource),sources,submissionsEnabled:String(cfg.allowSubmissions).toUpperCase()!=='FALSE'};
}
function savePlan(record){
 // 教師用の集計シートで扱うため、氏名やメールアドレスは収集しません。
 const data=getAppData(), cfg=getConfig_();
 if(String(cfg.allowSubmissions).toUpperCase()==='FALSE')throw new Error('現在、送信は受け付けていません。');
 if(!record || typeof record!=='object')throw new Error('送信内容が不正です。');
 const classCode=plain_(record.classCode,24),studentCode=plain_(record.studentCode,24);
 if(!/^[\w\-ぁ-んァ-ヶ一-龠]{1,24}$/.test(classCode)||!/^[\w\-ぁ-んァ-ヶ一-龠]{1,24}$/.test(studentCode)) throw new Error('クラスコード・生徒番号は1〜24文字で入力してください。');
 if(String(cfg.classPassphrase||'') && String(record.passphrase||'')!==String(cfg.classPassphrase))throw new Error('クラス用の合言葉を確認してください。');
 if(!['初案','改善案'].includes(record.stage))throw new Error('保存段階が不正です。');
 const ids=data.sources.map(s=>String(s.id));
 const plan=record.plan;
 if(!plan||typeof plan!=='object'||Array.isArray(plan)||Object.keys(plan).length!==ids.length)throw new Error('発電計画の項目数が正しくありません。');
 const values=ids.map(k=>Number(plan[k]));
 if(values.some(n=>!Number.isInteger(n)||n<0||n>100))throw new Error('割合は0〜100の整数で入力してください。');
 const total=values.reduce((a,b)=>a+b,0);
 if(total!==100)throw new Error('合計を100%にしてください。');
 const texts=['reason','comparison','constraints','disaster','revision'].map(k=>plain_(record[k],1800));
 if(texts[0].trim().length<10)throw new Error('選択理由を10文字以上で記入してください。');
 const initialPlan=record.initialPlan&&typeof record.initialPlan==='object'&&!Array.isArray(record.initialPlan)?record.initialPlan:null;
 if(record.stage==='改善案'&&!initialPlan)throw new Error('改善案には初案のデータが必要です。');
 if(initialPlan){if(ids.some(k=>!Number.isInteger(Number(initialPlan[k]))||Number(initialPlan[k])<0||Number(initialPlan[k])>100)||ids.reduce((a,k)=>a+Number(initialPlan[k]),0)!==100)throw new Error('初案の割合を確認してください。');}
 const more=['changeReason','disasterCountermeasure','outageImpact'].map(k=>plain_(record[k],1800));
 const row=[new Date(),classCode,studentCode,record.stage,JSON.stringify(plan),...texts,data.demandLabel,total,initialPlan?JSON.stringify(initialPlan):'',...more];
 upgradeRecords_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.RECORDS));
 const lock=LockService.getScriptLock();lock.waitLock(10000);
 try{SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.RECORDS).appendRow(row);}
 finally{lock.releaseLock();}
 return {ok:true,message:'計画を記録しました。',stage:record.stage};
}
function getConfig_(){const rows=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG).getDataRange().getValues();return Object.fromEntries(rows.slice(1).map(r=>[String(r[0]),r[1]]));}
function plain_(v,max){return String(v==null?'':v).replace(/[\u0000-\u001f]/g,' ').slice(0,max).replace(/^[=+@\-\t\r]/,m=>"'"+m);}
