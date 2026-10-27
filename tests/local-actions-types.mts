import {validateLocalActionDraft,LOCAL_ACTION_CATALOG,type LocalActionDescriptor,type LocalActionList} from '../core/local-actions.mjs';
import {LocalActions} from '../adapters/local-actions.mjs';
const descriptors:LocalActionDescriptor[]=[{schema:1,command:'freeze',parameters:{sheet:'S',rows:0,columns:0}},{schema:1,command:'formatRange',parameters:{sheet:'S',range:'A1',format:{bold:false,numberFormat:'date'}}},{schema:1,command:'recalculate',parameters:{}},{schema:1,command:'sortRows',parameters:{sheet:'S',range:'A1:B3',keys:[{column:'B',direction:'desc'}]}}];
async function storeContract(filename:string){const store=await LocalActions.open(filename);const list:LocalActionList=await store.list();await store.create(validateLocalActionDraft({name:'Action',descriptor:descriptors[0]}),list.revision);const entry=list.entries[0];if(entry){await store.update(entry.id,{name:entry.name,descriptor:entry.descriptor},list.revision);await store.remove(entry.id,list.revision);}}
void storeContract;void LOCAL_ACTION_CATALOG;
// @ts-expect-error Source style IDs are not portable saved action parameters.
const sourceStyle:LocalActionDescriptor={schema:1,command:'formatRange',parameters:{sheet:'S',range:'A1',format:{styleId:1}}};
// @ts-expect-error Recalculation contains no caller path or source arguments.
const path:LocalActionDescriptor={schema:1,command:'recalculate',parameters:{filename:'/tmp/a.xlsx'}};
// @ts-expect-error Author identity cannot be forged by a draft.
const forgedDraft:Parameters<LocalActions['create']>[0]={name:'Action',descriptor:descriptors[0],author:'admin'};
// @ts-expect-error Registry revisions are numeric, not workbook SHA strings.
const registryRevision:Parameters<LocalActions['remove']>[1]='abc123';
void sourceStyle;void path;void forgedDraft;void registryRevision;
