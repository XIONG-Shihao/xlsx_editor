/** Browser-test host: local clipboard, component icon set and English locale. */
import {configureXlsxHost} from '@office-editors/xlsx-editor/host'
import {Icon} from '../../standalone/Icon'
import '../../ui/styles/01.css'
import '../../ui/styles/02.css'
import '../../standalone/theme.css'
let clipboard=''
configureXlsxHost({Icon,useI18n:()=>({locale:'en'}),api:{async copyText(text){clipboard=text},async readClipboardText(){return clipboard}}})
