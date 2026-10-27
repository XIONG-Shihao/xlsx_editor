import { validateConditionalRules, validateValidationRules, containsFormula } from './rules-model.mjs';
import { refusal } from './refusal-codes.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const ruleChildren = (node) => Array.from(node?.childNodes ?? []).filter((item) => item.nodeType === 1);
function attrs(node, allowed) {
  if (
    node.namespaceURI !== NS ||
    Array.from(node.attributes).some(
      (item) =>
        item.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
        (item.namespaceURI || !allowed.includes(item.localName))
    )
  )
    throw refusal('rules-protected', 'Extended ' + node.localName + ' rules are protected');
}
function children(node, names) {
  if (
    ruleChildren(node).some((item) => item.namespaceURI !== NS || !names.includes(item.localName)) ||
    Array.from(node.childNodes).some((item) => item.nodeType === 3 && item.data.trim())
  )
    throw refusal('rules-protected', 'Opaque rule children are protected');
}
function boolean(value, fallback = false) {
  if (!value) return fallback;
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  throw refusal('internal-invalid-field', 'Invalid rule boolean');
}
const formulas = (node) =>
  ruleChildren(node).map((child) => {
    attrs(child, []);
    if (ruleChildren(child).length) throw refusal('rules-protected', 'Nested rule formulas are protected');
    return child.textContent;
  });
export function differentialFill(styles, index) {
  const containers = ruleChildren(styles?.documentElement).filter((item) => item.localName === 'dxfs');
  if (containers.length !== 1 || !Number.isSafeInteger(index) || index < 0)
    throw refusal('internal-missing-field', 'Missing differential fill');
  const dxf = ruleChildren(containers[0])[index];
  if (!dxf || dxf.localName !== 'dxf') throw refusal('internal-missing-field', 'Missing differential fill');
  attrs(dxf, []);
  children(dxf, ['fill']);
  const fill = ruleChildren(dxf)[0];
  if (ruleChildren(dxf).length !== 1) throw refusal('rules-protected', 'Advanced differential styles are protected');
  attrs(fill, []);
  children(fill, ['patternFill']);
  const pattern = ruleChildren(fill)[0];
  if (ruleChildren(fill).length !== 1) throw refusal('rules-protected', 'Advanced differential styles are protected');
  attrs(pattern, ['patternType']);
  children(pattern, ['fgColor', 'bgColor']);
  if (pattern.getAttribute('patternType') !== 'solid')
    throw refusal('rules-unsupported', 'Only solid fill rules are qualified');
  const fg = ruleChildren(pattern).filter((item) => item.localName === 'fgColor'),
    bg = ruleChildren(pattern).filter((item) => item.localName === 'bgColor');
  if (fg.length !== 1 || bg.length > 1) throw refusal('rules-protected', 'Opaque differential colors are protected');
  attrs(fg[0], ['rgb']);
  if (ruleChildren(fg[0]).length || !/^FF[0-9a-f]{6}$/i.test(fg[0].getAttribute('rgb')))
    throw refusal('rules-protected', 'Opaque differential color is protected');
  if (bg[0]) {
    const same = bg[0].hasAttribute('rgb');
    attrs(bg[0], [same ? 'rgb' : 'indexed']);
    if (
      ruleChildren(bg[0]).length ||
      (same
        ? bg[0].getAttribute('rgb').toUpperCase() !== fg[0].getAttribute('rgb').toUpperCase()
        : bg[0].getAttribute('indexed') !== '64')
    )
      throw refusal('rules-protected', 'Opaque background is protected');
  }
  return '#' + fg[0].getAttribute('rgb').slice(2).toUpperCase();
}
export function readConditionalRules(sheetDocument, stylesDocument) {
  const result = [],
    priorities = new Set();
  for (const container of ruleChildren(sheetDocument.documentElement).filter(
    (node) => node.localName === 'conditionalFormatting'
  )) {
    attrs(container, ['sqref']);
    children(container, ['cfRule']);
    for (const node of ruleChildren(container)) {
      attrs(node, ['type', 'operator', 'dxfId', 'priority', 'stopIfTrue', 'text']);
      children(node, ['formula']);
      const priority = Number(node.getAttribute('priority'));
      if (!Number.isSafeInteger(priority) || priority < 1 || priorities.has(priority))
        throw refusal('rules-ambiguous', 'Ambiguous conditional priority is protected');
      priorities.add(priority);
      const type = node.getAttribute('type'),
        operator = node.getAttribute('operator'),
        values = formulas(node);
      const rule = {
        range: container.getAttribute('sqref'),
        kind:
          type === 'duplicateValues'
            ? 'duplicateValues'
            : type === 'containsText'
              ? 'textContains'
              : type === 'cellIs'
                ? operator
                : null,
        fillColor: differentialFill(stylesDocument, Number(node.getAttribute('dxfId'))),
        stopIfTrue: boolean(node.getAttribute('stopIfTrue')),
      };
      if (type === 'duplicateValues') {
        if (values.length || operator || node.hasAttribute('text'))
          throw refusal('rules-protected', 'Extended duplicate rule is protected');
      } else if (type === 'containsText') {
        rule.text = node.getAttribute('text');
        if (operator !== 'containsText' || values.length !== 1 || values[0] !== containsFormula(rule.range, rule.text))
          throw refusal('rules-protected', 'Unqualified text formula is protected');
      } else if (type === 'cellIs') {
        if (
          node.hasAttribute('text') ||
          values.some((value) => !/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?$/.test(value))
        )
          throw refusal('rules-protected', 'Dynamic conditional formula is protected');
        rule.values = values.map(Number);
      } else throw refusal('rules-protected', 'Conditional rule type is protected');
      result.push({ ...validateConditionalRules([rule])[0], priority });
    }
  }
  result.sort((a, b) => a.priority - b.priority);
  return validateConditionalRules(result.map(({ priority, ...rule }) => rule));
}
export function readValidationRules(sheetDocument) {
  const containers = ruleChildren(sheetDocument.documentElement).filter((node) => node.localName === 'dataValidations');
  if (containers.length > 1) throw refusal('rules-ambiguous', 'Ambiguous validations are protected');
  if (!containers.length) return [];
  const container = containers[0];
  attrs(container, ['count']);
  children(container, ['dataValidation']);
  if (container.hasAttribute('count') && Number(container.getAttribute('count')) !== ruleChildren(container).length)
    throw refusal('internal-invalid-field', 'Invalid validation count');
  const rules = ruleChildren(container).map((node) => {
    attrs(node, [
      'type',
      'operator',
      'allowBlank',
      'showDropDown',
      'showInputMessage',
      'showErrorMessage',
      'errorStyle',
      'sqref',
      'promptTitle',
      'prompt',
      'errorTitle',
      'error',
    ]);
    children(node, ['formula1', 'formula2']);
    if (
      (node.getAttribute('errorStyle') || 'stop') !== 'stop' ||
      !boolean(node.getAttribute('showErrorMessage'), false)
    )
      throw refusal('rules-protected', 'Warning/disabled validation is protected');
    const type = node.getAttribute('type'),
      kind = type === 'whole' ? 'wholeNumber' : type,
      rule = {
        range: node.getAttribute('sqref'),
        kind,
        allowBlank: boolean(node.getAttribute('allowBlank')),
        showDropdown: kind === 'list' && !boolean(node.getAttribute('showDropDown')),
        promptTitle: node.getAttribute('promptTitle') || '',
        prompt: node.getAttribute('prompt') || '',
        errorTitle: node.getAttribute('errorTitle') || '',
        error: node.getAttribute('error') || '',
      },
      values = formulas(node);
    if (
      ruleChildren(node)
        .map((child) => child.localName)
        .join(',') !== (kind === 'list' ? 'formula1' : 'formula1,formula2')
    )
      throw refusal('rules-ambiguous', 'Ambiguous validation formulas are protected');
    if (
      boolean(node.getAttribute('showInputMessage'), true) !== !!(rule.promptTitle || rule.prompt) &&
      node.hasAttribute('showInputMessage')
    )
      throw refusal('rules-protected', 'Disabled prompt is protected');
    if (kind === 'list') {
      if (node.hasAttribute('operator') || values.length !== 1 || !/^"[^"]+"$/.test(values[0]))
        throw refusal('rules-protected', 'Dynamic validation lists are protected');
      rule.values = values[0].slice(1, -1).split(',');
    } else {
      if (
        (node.getAttribute('operator') || 'between') !== 'between' ||
        values.length !== 2 ||
        values.some((value) => !/^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?$/.test(value))
      )
        throw refusal('rules-limit', 'Dynamic validation bounds are protected');
      rule.minimum = Number(values[0]);
      rule.maximum = Number(values[1]);
    }
    return rule;
  });
  return validateValidationRules(rules);
}
export const ruleXmlText = (value) =>
  String(value).replace(
    /[&<>"\r\n\t]/g,
    (char) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\r': '&#13;', '\n': '&#10;', '\t': '&#9;' })[char]
  );
export function conditionalRulesXml(rules, prefix = '', firstDxf = 0) {
  const tag = (name) => prefix + name;
  return rules
    .map((rule, index) => {
      const type =
          rule.kind === 'duplicateValues'
            ? 'duplicateValues'
            : rule.kind === 'textContains'
              ? 'containsText'
              : 'cellIs',
        attributes = {
          type,
          ...(type === 'cellIs'
            ? { operator: rule.kind }
            : type === 'containsText'
              ? { operator: 'containsText', text: rule.text }
              : {}),
          dxfId: firstDxf + index,
          priority: index + 1,
          stopIfTrue: rule.stopIfTrue ? '1' : '0',
        },
        values = type === 'containsText' ? [containsFormula(rule.range, rule.text)] : (rule.values ?? []);
      return (
        '<' +
        tag('conditionalFormatting') +
        ' sqref="' +
        rule.range +
        '"><' +
        tag('cfRule') +
        Object.entries(attributes)
          .map(([key, value]) => ' ' + key + '="' + ruleXmlText(value) + '"')
          .join('') +
        '>' +
        values.map((value) => '<' + tag('formula') + '>' + ruleXmlText(value) + '</' + tag('formula') + '>').join('') +
        '</' +
        tag('cfRule') +
        '></' +
        tag('conditionalFormatting') +
        '>'
      );
    })
    .join('');
}
export function differentialRulesXml(rules, prefix = '') {
  const tag = (name) => prefix + name;
  return rules
    .map(
      (rule) =>
        '<' +
        tag('dxf') +
        '><' +
        tag('fill') +
        '><' +
        tag('patternFill') +
        ' patternType="solid"><' +
        tag('fgColor') +
        ' rgb="FF' +
        rule.fillColor.slice(1) +
        '"/><' +
        tag('bgColor') +
        ' indexed="64"/></' +
        tag('patternFill') +
        '></' +
        tag('fill') +
        '></' +
        tag('dxf') +
        '>'
    )
    .join('');
}
export function validationRulesXml(rules, prefix = '') {
  const tag = (name) => prefix + name;
  if (!rules.length) return '';
  return (
    '<' +
    tag('dataValidations') +
    ' count="' +
    rules.length +
    '">' +
    rules
      .map((rule) => {
        const attributes = {
            type: rule.kind === 'wholeNumber' ? 'whole' : rule.kind,
            ...(rule.kind === 'list' ? {} : { operator: 'between' }),
            allowBlank: rule.allowBlank ? '1' : '0',
            showDropDown: rule.kind === 'list' && !rule.showDropdown ? '1' : '0',
            showInputMessage: rule.prompt || rule.promptTitle ? '1' : '0',
            showErrorMessage: '1',
            errorStyle: 'stop',
            sqref: rule.range,
            promptTitle: rule.promptTitle,
            prompt: rule.prompt,
            errorTitle: rule.errorTitle,
            error: rule.error,
          },
          values = rule.kind === 'list' ? ['"' + rule.values.join(',') + '"'] : [rule.minimum, rule.maximum];
        return (
          '<' +
          tag('dataValidation') +
          Object.entries(attributes)
            .map(([key, value]) => ' ' + key + '="' + ruleXmlText(value) + '"')
            .join('') +
          '>' +
          values
            .map(
              (value, index) =>
                '<' +
                tag('formula' + (index + 1)) +
                '>' +
                ruleXmlText(value) +
                '</' +
                tag('formula' + (index + 1)) +
                '>'
            )
            .join('') +
          '</' +
          tag('dataValidation') +
          '>'
        );
      })
      .join('') +
    '</' +
    tag('dataValidations') +
    '>'
  );
}
