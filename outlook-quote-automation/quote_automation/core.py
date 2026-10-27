from __future__ import annotations
import dataclasses
import html
import json
import re
from datetime import datetime, timezone, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from html.parser import HTMLParser
from pathlib import Path

KST = timezone(timedelta(hours=9))

def now():
    return datetime.now(KST).isoformat(timespec='seconds')

class TextParser(HTMLParser):
    def __init__(self):
        super().__init__(); self.parts = []; self.hidden = 0
    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'): self.hidden += 1
        if tag in ('br', 'p', 'div', 'tr', 'li'): self.parts.append('\n')
    def handle_endtag(self, tag):
        if tag in ('script', 'style'): self.hidden = max(0, self.hidden - 1)
        if tag in ('p', 'div', 'tr', 'li'): self.parts.append('\n')
    def handle_data(self, data):
        if not self.hidden: self.parts.append(data)

def plain(message):
    body = message.get('body', {})
    if body.get('contentType', '').lower() == 'html':
        p = TextParser(); p.feed(body.get('content', '')); text = ''.join(p.parts)
    else: text = body.get('content', '')
    # 인용된 이전 회신은 기본 추출 대상에서 제외합니다.
    return re.split(r'(?im)^\s*(?:-----Original Message-----|보낸 사람:|From:|On .+ wrote:)', text)[0][:24000]

@dataclasses.dataclass
class Item:
    part: str
    description: str
    quantity: int | None
    unit: int | None = None
    status: str = '확인필요'
    note: str = ''
    evidence: dict = dataclasses.field(default_factory=dict)

def valid_items(values):
    if not isinstance(values, list) or len(values) > 100:
        raise ValueError('품목은 최대 100개까지 처리할 수 있습니다.')
    items = []
    for v in values:
        q = v.get('quantity')
        if isinstance(q, bool) or not isinstance(q, int) or not 0 < q <= 1000000: q = None
        part = str(v.get('part') or '')[:120].strip()
        desc = str(v.get('description') or part or '품목 확인필요')[:500].strip()
        items.append(Item(part, desc, q))
    return items or [Item('', '품목·모델·수량 확인필요', None, note='메일에서 품목을 추출하지 못했습니다.')]

def extract_rules(text):
    """명확한 '모델명 / 설명 / 수량' 또는 '모델명 3개'만 확정 추출."""
    values = []
    for line in text.splitlines():
        line = line.strip()
        m = re.match(r'^(?:[-*•]\s*)?([\w][\w.+()/ -]{1,100}?)\s*[|/]\s*(.*?)\s*[|/]\s*(\d[\d,]*)\s*(?:개|대|EA|ea|set|세트)?$', line)
        if m:
            values.append({'part': m[1].strip(), 'description': m[2].strip(), 'quantity': int(m[3].replace(',', ''))}); continue
        m = re.match(r'^(?:[-*•]\s*)?(?:모델(?:명)?\s*[:：]\s*)?([A-Za-z0-9][\w.+()/\-]{1,80})\s*[,;:]?\s*(?:수량\s*[:：]?\s*)?(\d[\d,]*)\s*(?:개|대|EA|ea|세트)\s*$', line)
        if m:
            values.append({'part': m[1], 'description': m[1], 'quantity': int(m[2].replace(',', ''))})
    result = {'customer': '', 'contact': '', 'phone': '', 'items': valid_items(values), 'notes': ''}
    for key, label in [('customer','고객사|회사명'),('contact','담당자|담당'),('phone','연락처|전화')]:
        m = re.search(r'(?m)^\s*(?:'+label+r')\s*[:：]\s*(.+)$', text)
        if m: result[key] = m[1].strip()[:120]
    return result

def round_won(value):
    return int(Decimal(str(value)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))

def price_with_margin(cost, cfg):
    rate = cfg.get('margin_percent')
    if rate is None: raise ValueError('마진율 미설정')
    try:
        rate = Decimal(str(rate)) / 100
        cost = Decimal(str(cost))
    except InvalidOperation:
        raise ValueError('가격 또는 마진율은 숫자로 설정하세요.') from None
    if not rate.is_finite() or not cost.is_finite(): raise ValueError('유효하지 않은 가격 또는 마진율')
    if cost <= 0: raise ValueError('유효하지 않은 가격')
    if cfg.get('search_price_vat') == 'included': cost /= Decimal('1.1')
    elif cfg.get('search_price_vat') != 'excluded': raise ValueError('검색 가격의 부가세 포함 여부 확인필요')
    mode = cfg.get('margin_mode', 'markup')
    if mode == 'markup':
        if not 0 <= rate <= 10: raise ValueError('가산율 범위: 0~1000%')
        result = cost * (1 + rate)
    elif mode == 'gross_margin':
        if not 0 <= rate < 1: raise ValueError('매출이익률 범위: 0~100% 미만')
        result = cost / (1 - rate)
    else: raise ValueError('margin_mode는 markup 또는 gross_margin이어야 합니다.')
    return round_won(result)

def totals(items):
    amounts = [round_won(x.quantity * x.unit) if x.quantity is not None and x.unit is not None else None for x in items]
    complete = bool(items) and all(x is not None for x in amounts)
    supply = sum(amounts) if complete else None
    vat = round_won(Decimal(supply) * Decimal('.1')) if complete else None
    return {'amounts': amounts, 'supply': supply, 'vat': vat, 'total': supply + vat if complete else None}

def recipient(message):
    candidates = message.get('replyTo') or [message.get('from', {})]
    if len(candidates) != 1:
        raise ValueError('회신 주소가 여러 개여서 자동으로 수신자를 정할 수 없습니다.')
    address = candidates[0].get('emailAddress', {})
    email = address.get('address', '')
    if not re.fullmatch(r'[^\s<>@]+@[^\s<>@]+\.[^\s<>@]+', email):
        raise ValueError('유효한 문의자 회신 주소를 확인할 수 없습니다.')
    return {'emailAddress': {'address': email, 'name': address.get('name', '')}}

def compose(message, quote, items, signature):
    unresolved = any(x.unit is None or x.quantity is None for x in items)
    # 재고, 납기, 최저가의 절대성, 계약 조건은 약속하지 않습니다.
    subject = re.sub(r'^(?:(?:RE|FW|FWD)\s*:\s*)+', '', message.get('subject', ''), flags=re.I).strip()
    subject = ('[견적 회신 / 확인필요] ' if unresolved else '[견적 회신] ') + subject[:180]
    esc = html.escape
    rows = ''.join(f'<tr><td>{esc(x.part or x.description)}</td><td>{x.quantity if x.quantity is not None else "확인필요"}</td><td>{format(x.unit, ",")+"원" if x.unit is not None else "확인필요"}</td></tr>' for x in items)
    total = totals(items)['total']
    notes = list(dict.fromkeys(x.note for x in items if x.note))
    body = f'''<p>안녕하세요. {esc(quote.get('contact') or '담당자')}님,</p>
<p>견적 문의해 주셔서 감사합니다. 요청하신 품목의 검토용 견적서를 첨부드립니다.</p>
<table border="1" cellpadding="8" cellspacing="0"><tr><th>모델 / 품목</th><th>수량</th><th>공급단가 · VAT 별도</th></tr>{rows}</table>
<p>{'부가세 포함 합계: '+format(total, ',')+'원' if total is not None else '확인되지 않은 품목 또는 수량이 있어 전체 견적금액은 확인필요입니다.'}</p>
<p>단가는 검색 시점에 확인한 공개 판매가를 바탕으로 산정한 참고 견적입니다. 검색 범위 내 최저가이며, 재고·배송비·납기·판매조건은 최종 확인이 필요합니다.</p>
{''.join('<p>• '+esc(n)+'</p>' for n in notes)}
<p>‘확인필요’로 표시된 항목은 모델·수량·조건을 확인한 후 확정 견적을 안내드리겠습니다.<br>추가 요청사항이 있으시면 회신 부탁드립니다. 감사합니다.</p>
<p>{esc(signature).replace(chr(10), '<br>')}</p>'''
    return {'subject': subject, 'body': {'contentType': 'HTML', 'content': body}, 'toRecipients': [recipient(message)], 'ccRecipients': [], 'bccRecipients': []}

def safe_text(value):
    # 외부 텍스트가 엑셀 수식으로 실행되지 않도록 문자열로 기록합니다.
    text = str(value or '')
    return "'" + text if text.startswith(('=', '+', '-', '@')) else text
