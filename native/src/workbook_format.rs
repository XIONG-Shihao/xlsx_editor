//! Neutral component-owned identifiers emitted for new workbook data.
//! Current component identities; retired app namespaces are unsupported.
macro_rules! workbook_format {
 (PRINT_DEFINITION_V1)=>{"workbook-print-definition-v1"};
 (RULES_V1)=>{"workbook-rules-v1"};
    (SCHEMA_NOTE_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.note-calculation-transfer.v1" };
    (SCHEMA_BACKGROUND_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.background-calculation-transfer.v1" };
    (SCHEMA_PAGE_THEME_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.page-theme-calculation-transfer.v1" };
    (SCHEMA_HYPERLINK_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.hyperlink-calculation-transfer.v1" };
    (SCHEMA_CLEAR_HYPERLINKS_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.clear-hyperlinks-calculation-transfer.v1" };
    (SCHEMA_PUBLICATION_CALCULATION_TRANSFER_V1) => { "workbook.xlsx.publication-calculation-transfer.v1" };
    (SCHEMA_PUBLICATION_CALCULATION_RESTORE_V1) => { "workbook.xlsx.publication-calculation-restore.v1" };
    (URN_OWNED_FILTER_V1) => { "urn:workbook:owned-filter:v1" };
    (SCHEMA_CUSTOM_FORMAT_VECTORS_V1) => { "workbook.xlsx.custom-format-vectors.v1" };
    (SCHEMA_STAGED_PASTE_V2) => { "workbook.xlsx.staged-paste.v2" };
    (SCHEMA_STAGED_PASTE_V1) => { "workbook.xlsx.staged-paste.v1" };
    (SCHEMA_BACKGROUND_WRITER_V1) => { "workbook.xlsx.background-writer.v1" };
    (XL_MEDIA_BACKGROUND_ID_PNG) => { "xl/media/workbook-background-{}-{id}.png" };
    (SCHEMA_SHEET_BACKGROUND_V1) => { "workbook.xlsx.sheet-background.v1" };
    (BACKGROUND_SUFFIX) => { "workbookBackground{suffix}" };
    (SCHEMA_BOOLEAN_CELL_ADMISSION_V1) => { "workbook.xlsx.boolean-cell-admission.v1" };
    (WORKBOOK_INDEX_OWNER_V1) => { "workbook.workbook.index-owner.v1" };
    (SCHEMA_CLEAR_HYPERLINKS_WRITER_V1) => { "workbook.xlsx.clear-hyperlinks-writer.v1" };
    (SCHEMA_HYPERLINK_WRITER_V1) => { "workbook.xlsx.hyperlink-writer.v1" };
    (XML_B441F277FE) => { "r:id=\"workbookHyperlink1\"" };
    (HYPERLINK1) => { "workbookHyperlink1" };
    (HYPERLINK_INDEX) => { "workbookHyperlink{index}" };
    (SCHEMA_HYPERLINK_DEFINITION_V1) => { "workbook.xlsx.hyperlink-definition.v1" };
    (ID) => { "workbook:{id}" };
    (XL_MEDIA_IMAGE) => { "xl/media/workbook-image-" };
    (XML_6C801357F7) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{}\" xmlns:a=\"{}\"><xdr:from><xdr:col>{col}</xdr:col",
        "><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowO",
        "ff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}\"/><xdr:sp><xdr:nvSpPr><xdr:cNvP",
        "r id=\"{number}\" name=\"workbook:{id}\" descr=\"{}\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr",
        ":spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></a:xfrm><a",
        ":custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"{width",
        "}\" b=\"{height}\"/><a:pathLst><a:path w=\"{width}\" h=\"{height}\" fill=\"none\" stroke=",
        "\"1\" extrusionOk=\"0\">{steps}</a:path></a:pathLst></a:custGeom><a:noFill/><a:ln w=",
        "\"{pen_width}\" cap=\"{cap}\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:srgbClr val=\"{co",
        "lor}\"><a:alpha val=\"{opacity}\"/></a:srgbClr></a:solidFill><a:prstDash val=\"solid",
        "\"/><a:round/><a:headEnd type=\"none\"/><a:tailEnd type=\"none\"/></a:ln></xdr:spPr><",
        "/xdr:sp><xdr:clientData/></xdr:oneCellAnchor>",
    ) };
    (XML_3E43FDA91C) => { "id=\"1\" name=\"workbook:" };
    (XML_870DB7DE39) => { "id=\"4\" name=\"workbook:" };
    (URN_MODEL3D_SAMPLES_V1) => { "urn:workbook:model3d:samples:v1" };
    (MODEL3D) => { "workbookModel3d" };
    (CLOSED_MODEL3D_SAMPLES_V1) => { "Workbook closed model3d samples v1" };
    (MODEL3D_D80344) => { "workbookModel3d_{}" };
    (XL_MEDIA_MODEL3D_ID) => { "xl/media/workbook-model3d-{id}-" };
    (XML_4C3CF49590) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{d}\" xmlns:a=\"{a}\" xmlns:r=\"{r}\" xmlns:am3d=\"{M}\" ",
        "xmlns:workbook=\"{PRIVATE}\"><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{dx}</x",
        "dr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowOff></xdr:from><xdr:e",
        "xt cx=\"{w}\" cy=\"{h}\"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{nu",
        "mber}\" name=\"workbook:{id}\" descr=\"{}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicF",
        "ramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></xdr:xfrm><a:gra",
        "phic><a:graphicData uri=\"{M}\"><am3d:model3d r:embed=\"{rid}\"><am3d:spPr/><am3d:ca",
        "mera><am3d:pos x=\"0\" y=\"0\" z=\"180000000\"/><am3d:up dx=\"0\" dy=\"1\" dz=\"0\"/><am3d:l",
        "ookAt x=\"0\" y=\"0\" z=\"0\"/><am3d:orthographic><am3d:sz n=\"{ortho_n}\" d=\"{ortho_d}\"",
        "/></am3d:orthographic></am3d:camera><am3d:trans><am3d:meterPerModelUnit n=\"1\" d=",
        "\"1\"/></am3d:trans><am3d:extLst><a:ext uri=\"{PRIVATE}\"><workbook:sample sample=\"{",
        "sample}\" color=\"{color}\" x=\"{x}\" y=\"{y}\" z=\"{z}\"/></a:ext></am3d:extLst><am3d:wi",
        "nViewport/></am3d:model3d></a:graphicData></a:graphic></xdr:graphicFrame><xdr:cl",
        "ientData/></xdr:oneCellAnchor>",
    ) };
    (XML_BC0D1F4093) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{}\" xmlns:a=\"{}\"><xdr:from><xdr:col>{col}</xdr:col",
        "><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowO",
        "ff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}\"/><xdr:sp><xdr:nvSpPr><xdr:cNvP",
        "r id=\"{number}\" name=\"workbook:{id}\" descr=\"{}\"/><xdr:cNvSpPr txBox=\"0\"/></xdr:n",
        "vSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/><",
        "/a:xfrm>{geometry}{}<a:ln w=\"12700\"><a:noFill/></a:ln></xdr:spPr></xdr:sp><xdr:c",
        "lientData/></xdr:oneCellAnchor>",
    ) };
    (XML_A5FAA3BD99) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{}\" xmlns:a=\"{}\"><xdr:from><xdr:col>{col}</xdr:col",
        "><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowO",
        "ff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}\"/><xdr:sp><xdr:nvSpPr><xdr:cNvP",
        "r id=\"{number}\" name=\"workbook:{id}\" descr=\"{}\"/><xdr:cNvSpPr txBox=\"1\"/></xdr:n",
        "vSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/><",
        "/a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700",
        "\"><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr fromWordArt=\"1\" wrap=\"squar",
        "e\" anchor=\"ctr\" lIns=\"91440\" rIns=\"91440\" tIns=\"45720\" bIns=\"45720\"/><a:lstStyle",
        "/>{paragraphs}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>",
    ) };
    (XML_853DF0B88B) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{}\" xmlns:a=\"{}\"><xdr:from><xdr:col>{col}</xdr:col",
        "><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{dy}</xdr:rowO",
        "ff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}\"/><xdr:sp><xdr:nvSpPr><xdr:cNvP",
        "r id=\"{number}\" name=\"workbook:{id}\" descr=\"{}\"/><xdr:cNvSpPr txBox=\"{txbox}\"/><",
        "/xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{heigh",
        "t}\"/></a:xfrm><a:prstGeom prst=\"{geometry}\"><a:avLst/></a:prstGeom>{}<a:ln w=\"12",
        "700\">{}</a:ln></xdr:spPr><xdr:txBody><a:bodyPr wrap=\"square\" anchor=\"ctr\" lIns=\"",
        "91440\" rIns=\"91440\" tIns=\"45720\" bIns=\"45720\"/><a:lstStyle/>{paragraphs}</xdr:tx",
        "Body></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>",
    ) };
    (XL_DIAGRAMS_SMARTART_ID_ROLE_XML) => { "xl/diagrams/workbook-smartart-{id}-{role}.xml" };
    (XL_DIAGRAMS_RELS_SMARTART_ID_DATA_XML_RELS) => { "xl/diagrams/_rels/workbook-smartart-{id}-data.xml.rels" };
    (URN_SMARTART_LAYOUT_KIND) => { "urn:workbook:smartart:layout:{kind}" };
    (XML_F6A039DC2D) => { concat!(
        "<dgm:dataModel xmlns:dgm=\"{G}\" xmlns:a=\"{A}\" xmlns:dsp=\"{DSP}\"><dgm:ptLst>{point",
        "s}</dgm:ptLst><dgm:cxnLst>{connections}</dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:ex",
        "tLst><a:ext uri=\"{DSP}\"><dsp:dataModelExt relId=\"workbookDiagramDrawing\" minVer=",
        "\"{G}\"/></a:ext></dgm:extLst></dgm:dataModel>",
    ) };
    (XML_2FD003E147) => { concat!(
        "<dgm:layoutDef xmlns:dgm=\"{G}\" xmlns:a=\"{A}\" uniqueId=\"{uid}\" minVer=\"12.0\"><dgm",
        ":title val=\"Workbook {label}\"/><dgm:desc val=\"Repository-authored bounded sample",
        "\"/><dgm:catLst><dgm:cat type=\"{kind}\" pri=\"1\"/></dgm:catLst><dgm:layoutNode name",
        "=\"diagram\"><dgm:alg type=\"{alg}\"/><dgm:presOf axis=\"self\"/><dgm:forEach axis=\"{a",
        "xis}\" ptType=\"node\" hideLastTrans=\"{hide}\"><dgm:layoutNode name=\"node\" styleLbl=",
        "\"node1\"><dgm:alg type=\"sp\"/><dgm:presOf axis=\"self\" ptType=\"node\"/><dgm:shape ty",
        "pe=\"roundRect\"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach><dg",
        "m:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:layoutDef>",
    ) };
    (XML_AE2E00CE8C) => { concat!(
        "<dgm:styleDef xmlns:dgm=\"{G}\" xmlns:a=\"{A}\" uniqueId=\"urn:workbook:smartart:styl",
        "e:flat\" minVer=\"12.0\"><dgm:title val=\"Workbook Flat\"/><dgm:desc val=\"Flat sample",
        "\"/><dgm:catLst><dgm:cat type=\"simple\" pri=\"1\"/></dgm:catLst><dgm:styleLbl name=\"",
        "node1\"><dgm:txPr/><dgm:style>{refs}</dgm:style></dgm:styleLbl></dgm:styleDef>",
    ) };
    (XML_7BF15A0DC6) => { concat!(
        "<dgm:colorsDef xmlns:dgm=\"{G}\" xmlns:a=\"{A}\" uniqueId=\"urn:workbook:smartart:col",
        "ors:blue\" minVer=\"12.0\"><dgm:title val=\"Workbook Blue\"/><dgm:desc val=\"Blue and ",
        "white sample\"/><dgm:catLst><dgm:cat type=\"mainScheme\" pri=\"1\"/></dgm:catLst><dgm",
        ":styleLbl name=\"node1\"><dgm:fillClrLst><a:srgbClr val=\"4472C4\"/></dgm:fillClrLst",
        "><dgm:linClrLst><a:srgbClr val=\"2F5597\"/></dgm:linClrLst><dgm:effectClrLst/><dgm",
        ":txLinClrLst/><dgm:txFillClrLst><a:srgbClr val=\"FFFFFF\"/></dgm:txFillClrLst><dgm",
        ":txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>",
    ) };
    (XML_5961861C5F) => { concat!(
        "<Relationships xmlns=\"{P}\"><Relationship Id=\"workbookDiagramDrawing\" Type=\"http:",
        "//schemas.microsoft.com/office/2007/relationships/diagramDrawing\" Target=\"workbo",
        "ok-smartart-{id}-drawing.xml\"/></Relationships>",
    ) };
    (DIAGRAM_COMPACT) => { "workbookDiagram_{compact}" };
    (XML_D9E50A0C14) => { concat!(
        "<xdr:oneCellAnchor xmlns:xdr=\"{D}\" xmlns:a=\"{A}\" xmlns:dgm=\"{G}\" xmlns:r=\"{R}\"><",
        "xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{dx}</xdr:colOff><xdr:row>{row}</xd",
        "r:row><xdr:rowOff>{dy}</xdr:rowOff></xdr:from><xdr:ext cx=\"{width}\" cy=\"{height}",
        "\"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{number}\" name=\"workbo",
        "ok:{id}\" descr=\"{}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a",
        ":off x=\"0\" y=\"0\"/><a:ext cx=\"{width}\" cy=\"{height}\"/></xdr:xfrm><a:graphic><a:gr",
        "aphicData uri=\"{G}\"><dgm:relIds r:dm=\"{prefix}_dm\" r:lo=\"{prefix}_lo\" r:qs=\"{pre",
        "fix}_qs\" r:cs=\"{prefix}_cs\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr",
        ":clientData/></xdr:oneCellAnchor>",
    ) };
    (URN_SMARTART_LAYOUT) => { "urn:workbook:smartart:layout:" };
    (SCHEMA_NOTE_WRITER_V1) => { "workbook.xlsx.note-writer.v1" };
    (XL_COMMENTS_NOTES_XML) => { "xl/comments/workbook-notes-{}.xml" };
    (XL_DRAWINGS_NOTES_VML) => { "xl/drawings/workbook-notes-{}.vml" };
    (NOTES) => { "workbookNotes" };
    (NOTESVML) => { "workbookNotesVml" };
    (XML_3EA451F026) => { concat!(
        "<Relationship xmlns=\"{}\" Id=\"workbookNotes\" Type=\"{COMMENTS_EDGE}\" Target=\"/{}\"/",
        "><Relationship xmlns=\"{}\" Id=\"workbookNotesVml\" Type=\"{VML_EDGE}\" Target=\"/{}\"/>",
    ) };
    (XML_EC5CB73564) => { "<legacyDrawing xmlns=\"{}\" xmlns:r=\"{}\" r:id=\"workbookNotesVml\"/>" };
    (SCHEMA_LEGACY_NOTES_V1) => { "workbook.xlsx.legacy-notes.v1" };
    (XL_PIVOTTABLES_PIVOT1_XML) => { "xl/pivotTables/workbookPivot1.xml" };
    (XL_PIVOTCACHE_CACHEDEFINITION1_XML) => { "xl/pivotCache/workbookCacheDefinition1.xml" };
    (XL_PIVOTCACHE_CACHERECORDS1_XML) => { "xl/pivotCache/workbookCacheRecords1.xml" };
    (PIVOTTABLES_PIVOT1_XML) => { "../pivotTables/workbookPivot1.xml" };
    (PIVOTCACHE_CACHEDEFINITION1_XML) => { "../pivotCache/workbookCacheDefinition1.xml" };
    (CACHERECORDS1_XML) => { "workbookCacheRecords1.xml" };
    (THEME_PAGE_THEME_XML) => { "theme/workbook-page-theme.xml" };
    (XL_THEME_PAGE_THEME_XML) => { "xl/theme/workbook-page-theme.xml" };
    (SCHEMA_PAGE_THEME_WRITER_V1) => { "workbook.xlsx.page-theme-writer.v1" };
    (PAGETHEME_N) => { "workbookPageTheme{n}" };
    (XML_BECA37B7DA) => { "<Relationship xmlns=\"{}\" Id=\"workbookPageTheme{n}\" Type=\"{THEME_EDGE}\" Target=\"theme/workbook-page-theme.xml\"/>" };
    (XL_CHARTS_CHART) => { "xl/charts/workbook-chart-" };
    (FORMAT_) => { "workbook:" };
    (XL_MEDIA_MODEL3D_ID_SHA_GLB) => { "xl/media/workbook-model3d-{id}-{sha}.glb" };
    (XL_MEDIA_IMAGE_ID_PNG) => { "xl/media/workbook-image-{id}-{}.png" };
    (XL_DRAWINGS_DRAWING_ID_XML) => { "xl/drawings/workbook-drawing-{id}.xml" };
    (XL_CHARTS_CHART_ID_XML) => { "xl/charts/workbook-chart-{id}.xml" };
    (SPILL_NUMBER) => { "workbookSpill{number}" };
    (XML_FFB1491B7F) => { concat!(
        "<Relationships xmlns=\"{REL}\"><Relationship Id=\"workbookSpill1\" Type=\"urn:origina",
        "l\" Target=\"opaque\"/></Relationships>",
    ) };
    (XML_9A9AA58ED2) => { "Id=\"workbookSpill2\"" };
    (XML_906A8FF106) => { "Id=\"workbookSpill5\"" };
    (FILTERSTATE) => { "workbook:filterState" };
    (XMLNS) => { "xmlns:workbook" };
    (SCHEMA_USED_RANGE_FACTS_V2) => { "workbook.xlsx.used-range-facts.v2" };
    (SCHEMA_USED_RANGE_FACTS_V1) => { "workbook.xlsx.used-range-facts.v1" };
    (SCHEMA_USED_RANGE_V1) => { "workbook.xlsx.used-range.v1" };
    (SCHEMA_USED_RANGE_EXPORT_NEEDED_V1) => { "workbook.xlsx.used-range-export-needed.v1" };
    (SCHEMA_USED_RANGE_REWRITE_SEAL_V1) => { "workbook.xlsx.used-range-rewrite-seal.v1" };
    (SCHEMA_USED_RANGE_TRANSFER_V1) => { "workbook.xlsx.used-range-transfer.v1" };
    (SCHEMA_USED_RANGE_SOURCE_SNAPSHOT_V1) => { "workbook.xlsx.used-range-source-snapshot.v1" };
}

macro_rules! format_matches {($value:expr,$key:ident)=>{{let value=&$value;*value==workbook_format!($key)}};}
macro_rules! format_prefix {($value:expr,$key:ident)=>{{$value.strip_prefix(workbook_format!($key))}};}
macro_rules! format_identity_matches {($value:expr,$key:ident,$($args:tt)*)=>{{$value==format!(workbook_format!($key),$($args)*)}};}
pub(crate) fn page_theme_matches(bytes:&[u8],canonical:&str)->bool{bytes==canonical.as_bytes()}
