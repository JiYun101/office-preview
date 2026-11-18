function element(tag, className, text) {
  const node = document.createElement(tag); node.className = className;
  if (text !== undefined) node.textContent = String(text);
  return node;
}

export function renderEmail(model, content) {
  const attachmentUrls = [];
  const card = element('article', 'op-email');
  card.append(element('div', 'op-email-label', '邮件预览'), element('h1', 'op-email-subject', model.subject || '（无主题）'));
  const headers = element('dl', 'op-email-headers');
  for (const [label, value] of [['发件人', model.from], ['收件人', model.to], ['抄送', model.cc], ['日期', model.date]]) {
    if (!value && ['抄送', '日期'].includes(label)) continue;
    headers.append(element('dt', '', label), element('dd', '', Array.isArray(value) ? value.join(', ') : value || '（未提供）'));
  }
  card.append(headers);
  const body = element('pre', 'op-email-body', model.text || '（没有可显示的文字正文）');
  body.setAttribute('aria-label', '邮件正文'); card.append(body);
  if (model.attachments?.length) {
    const section = element('section', 'op-email-attachments');
    section.append(element('h2', '', `附件清单（${model.attachments.length}）`));
    const list = element('ul', '');
    for (const attachment of model.attachments) {
      const item = element('li', '');
      const size = attachment.size || 0;
      const name = attachment.name || '未命名附件';
      item.append(element('strong', '', name), element('span', '', `${attachment.mime || 'application/octet-stream'} · ${size < 1024 ? `${size} B` : `${(size / 1024).toFixed(1)} KB`}`));
      if (attachment.bytes instanceof Uint8Array) {
        const link = element('a', '', '下载');
        const url = URL.createObjectURL(new Blob([attachment.bytes], { type: 'application/octet-stream' }));
        attachmentUrls.push(url);
        link.href = url;
        link.download = name;
        link.setAttribute('aria-label', `下载 ${name}`);
        item.append(link);
      }
      list.append(item);
    }
    section.append(list); card.append(section);
  }
  content.append(card);
  return { fitWidth: 850, dispose() { for (const url of attachmentUrls.splice(0)) URL.revokeObjectURL(url); } };
}
