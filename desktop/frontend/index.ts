import {initialize,invoke,subscribe} from './ports'
import {createStandaloneApi} from '../../standalone/bridge.mjs'
Object.defineProperty(window,'xlsxDesktop',{value:createStandaloneApi({invoke,subscribe}),configurable:false})
void initialize().then(()=>{document.getElementById('root')!.removeAttribute('role');return import('../../standalone/renderer.tsx')}).catch(error=>{
  console.error(error)
  const root=document.getElementById('root')!
  root.textContent='The spreadsheet service could not start. Restart the editor to try again.'
  root.setAttribute('role','alert')
})
