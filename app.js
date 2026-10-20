/* BibWord MVP: Office.js only; no account or external bibliography service. */
let library = {};
const $ = id => document.getElementById(id);
function status(message,error=false){$('status').textContent=message;$('status').className=error?'error':'';}
function cleaned(s){return String(s||'').replace(/[{}]/g,'').replace(/\\([&%_#])/g,'$1').replace(/\\(?:emph|textit|textbf)\s*/g,'').replace(/\\[a-zA-Z]+/g,'').replace(/\s+/g,' ').trim();}
function parseBib(src){
  // Stateful scanner: handles nested braces, quoted values, escaped delimiters and concatenation.
  const out={}; let i=0; const len=src.length;
  const white=()=>{while(i<len && /[\s,]/.test(src[i]))i++;};
  const group=()=>{let open=src[i++], close=open==='{'?'}':'"', depth=1, v='';
    while(i<len && depth){let c=src[i++];if(c==='\\'&&i<len){v+=c+src[i++];continue;}
      if(open==='{'&&c==='{')depth++;else if(c===close)depth--;
      if(depth)v+=c;
    }return v;
  };
  const val=()=>{let chunks=[];while(i<len){while(i<len&&/\s/.test(src[i]))i++;
      if(src[i]==='{'||src[i]==='"')chunks.push(group());
      else {let st=i;while(i<len&&!/[,}\s#)]/.test(src[i]))i++;if(i===st)break;chunks.push(src.slice(st,i));}
      while(i<len&&/\s/.test(src[i]))i++;if(src[i]==='#'){i++;continue;}break;
    }return chunks.join('').trim();};
  while(i<len){let at=src.indexOf('@',i);if(at<0)break;i=at+1;
    let st=i;while(i<len&&/[\w-]/.test(src[i]))i++;let type=src.slice(st,i).toLowerCase();
    while(i<len&&/\s/.test(src[i]))i++;let open=src[i];if(open!=='{'&&open!=='('){i++;continue;}i++;
    if(['comment','preamble','string'].includes(type)){
      let depth=1;while(i<len&&depth){let c=src[i++];if(c===open)depth++;else if(c===(open==='{'?'}':')'))depth--;}continue;
    }
    st=i;while(i<len&&!/[,}\)]/.test(src[i]))i++;let key=src.slice(st,i).trim();if(src[i]!==','){continue;}i++;
    const e={key,type};
    while(i<len){white();if(src[i]==='}'||src[i]===')'){i++;break;}
      st=i;while(i<len&&/[\w-]/.test(src[i]))i++;let field=src.slice(st,i).toLowerCase();
      if(!field){i++;continue;}while(i<len&&/\s/.test(src[i]))i++;if(src[i]!=='='){continue;}i++;e[field]=val();
    }if(key)out[key]=e;
  }return out;
}
function author(e){return cleaned(e.author||e.editor||'Unknown').replace(/\s+and\s+/gi,', ')}
function label(e){return cleaned(e.title||e.key)}
function format(e,n){let parts=[`[${n}] ${author(e)}.`, `“${label(e)}.”`];let venue=cleaned(e.journal||e.booktitle||e.publisher||'');if(venue)parts.push(venue+'.');if(e.year||e.date)parts.push(cleaned(e.year||e.date)+'.');if(e.volume)parts.push('Vol. '+cleaned(e.volume)+'.');if(e.pages)parts.push('pp. '+cleaned(e.pages)+'.');if(e.doi)parts.push('doi: '+cleaned(e.doi)+'.');return parts.join(' ')}
async function importText(content,source){
  const loaded=parseBib(content);
  const keys=Object.keys(loaded);
  if(!keys.length)throw Error('没有识别出 BibTeX 文献：请粘贴完整 @article{key,...} 条目。');
  const duplicates=keys.filter(k=>Object.prototype.hasOwnProperty.call(library,k));
  const previous=library;
  library={...library,...loaded};
  try{await save()}catch(e){library=previous;throw e;}
  render();status(`${source}成功：${keys.length} 条文献；新增 ${keys.length-duplicates.length} 条，覆盖 ${duplicates.length} 条。`);
}
function save(){return new Promise((resolve,reject)=>{Office.context.document.settings.set('bibword_library_v1',JSON.stringify(library));Office.context.document.settings.saveAsync(r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(new Error(r.error?.message||'保存失败')))});}
function render(){const q=$('search').value.trim().toLowerCase(),entries=Object.values(library).filter(e=>!q||[e.key,e.title,e.author,e.year].some(x=>String(x||'').toLowerCase().includes(q)));$('count').textContent=`已载入 ${Object.keys(library).length} 条文献`;const root=$('results');root.replaceChildren();for(const e of entries.slice(0,60)){let div=document.createElement('div');div.className='entry';let b=document.createElement('strong');b.textContent=label(e);let small=document.createElement('small');small.textContent=e.key+' · '+cleaned(e.year||e.date||'');let btn=document.createElement('button');btn.textContent='插入引用';btn.onclick=()=>insert(e.key);div.append(b,small,btn);root.append(div)}if(entries.length>60){const p=document.createElement('p');p.className='hint';p.textContent='仅显示前60项，请输入关键词缩小范围。';root.append(p)}}
async function insert(key){try{await Word.run(async ctx=>{let r=ctx.document.getSelection();const cc=r.insertContentControl();cc.tag='bibword:cite:'+key;cc.title='BibWord '+key;cc.appearance='Hidden';cc.insertText('[?]','Replace');await ctx.sync()});await refresh();status('已插入 '+key)}catch(e){status('插入失败: '+e.message,true)}}
async function getControls(ctx){let cc=ctx.document.body.contentControls;cc.load('items/tag');await ctx.sync();return cc.items}
async function replaceTokens(){try{let keys=Object.keys(library);let replaced=0;await Word.run(async ctx=>{let body=ctx.document.body;const hits=body.search('\\cite{*}',{matchWildcards:true});hits.load('items/text');await ctx.sync();for(const hit of hits.items){let m=hit.text.match(/^\\cite\{([^}]+)\}$/);if(!m)continue;let key=m[1].trim();if(!library[key])continue;let cc=hit.insertContentControl();cc.tag='bibword:cite:'+key;cc.title='BibWord '+key;cc.appearance='Hidden';cc.insertText('[?]','Replace');replaced++}await ctx.sync()});await refresh();status(`成功替换 ${replaced} 处引用。支持单 key 的 \\cite{key}。`)}catch(e){status('替换失败: '+e.message,true)}}
async function refresh(){try{let missing=[];await Word.run(async ctx=>{let controls=await getControls(ctx);let order=[];let bibs=[];for(let cc of controls){if(cc.tag?.startsWith('bibword:cite:')){let key=cc.tag.slice('bibword:cite:'.length);if(!order.includes(key))order.push(key)}else if(cc.tag==='bibword:bibliography')bibs.push(cc)}let n=new Map(order.map((key,i)=>[key,i+1]));for(let cc of controls){if(cc.tag?.startsWith('bibword:cite:')){let key=cc.tag.slice('bibword:cite:'.length);cc.insertText('['+n.get(key)+']','Replace')}}const lines=order.map((key,i)=>library[key]?format(library[key],i+1):(missing.push(key),`[${i+1}] ⚠ 未在当前文献库找到 ${key}`));for(let b of bibs)b.insertText('References\n'+lines.join('\n'),'Replace');await ctx.sync()});status('引用编号已更新。'+(missing.length?' 缺少: '+missing.join(', '):''),missing.length>0)}catch(e){status('更新失败: '+e.message,true)}}
async function bibliography(){try{let exists=false;await Word.run(async ctx=>{let controls=await getControls(ctx);exists=controls.some(x=>x.tag==='bibword:bibliography');if(!exists){let p=ctx.document.body.insertParagraph('References','End');let cc=p.insertContentControl();cc.tag='bibword:bibliography';cc.title='BibWord References';cc.appearance='Hidden';cc.insertText('References','Replace')}await ctx.sync()});await refresh();status(exists?'已更新现有参考文献表。':'已在文末生成参考文献表。')}catch(e){status('生成失败: '+e.message,true)}}
Office.onReady(info=>{if(info.host!==Office.HostType.Word){status('请在 Word 内打开此插件',true);return;}try{library=JSON.parse(Office.context.document.settings.get('bibword_library_v1')||'{}')}catch{library={}}render();status('Word 已连接。粘贴 BibTeX 或选择 .bib 文件即可导入。')});
$('file').addEventListener('change',async ev=>{let f=ev.target.files?.[0];if(!f)return;try{await importText(await f.text(),'文件导入')}catch(e){status(e.message,true)}finally{ev.target.value=''}});
$('pasteImport').onclick=async()=>{try{const v=$('paste').value.trim();if(!v)throw Error('请先粘贴 BibTeX 内容。');await importText(v,'粘贴导入')}catch(e){status(e.message,true)}};
$('clearPaste').onclick=()=>{$('paste').value='';status('已清空输入框，已导入的文献不受影响。')};
$('search').addEventListener('input',render);$('refresh').onclick=refresh;$('bibliography').onclick=bibliography;$('replace').onclick=replaceTokens;
