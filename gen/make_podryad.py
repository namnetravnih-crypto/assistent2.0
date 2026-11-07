# -*- coding: utf-8 -*-
import sys, glob, docx, copy
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
from docx.shared import Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
src, out = sys.argv[1], sys.argv[2]
d = docx.Document(src); P = d.paragraphs
def settext(p, text):
    runs = p.runs
    if not runs: p.add_run(text); return
    runs[0].text = text
    for r in runs[1:]: r._r.getparent().remove(r._r)
def drop(p): p._p.getparent().remove(p._p)
# запоминаем объекты заранее (индексы сдвигаются при удалении)
p0, p1, p2, p7 = P[0], P[1], P[2], P[7]
req_exec, req_cust = P[59:68], P[68:78]
sig_exec, sig_cust = P[83], P[87]
settext(p0, 'ДОГОВОР ПОДРЯДА №{{doc.number}}')
# город и дата
settext(p1, 'г. {{doc.city}}\t{{doc.date_long}}')
pPr = p1._p.get_or_add_pPr()
for t in pPr.findall(qn('w:tabs')): pPr.remove(t)
tabs = OxmlElement('w:tabs'); t = OxmlElement('w:tab'); t.set(qn('w:val'), 'right'); t.set(qn('w:pos'), '9071'); tabs.append(t); pPr.insert(0, tabs)
settext(p2, '{{executor.intro|Подрядчик}}, с одной стороны, и {{customer.intro|Заказчик}}, с другой стороны, именуемые в дальнейшем «Стороны», заключили настоящий договор (в дальнейшем «Договор») о нижеследующем:')
settext(p7, '2.1. Стоимость Работ по настоящему Договору составляет {{doc.amount}}, {{doc.vat_phrase}}. С момента подписания настоящего Договора указанная стоимость считается фиксированной и не может быть увеличена.')
def block(ps, role, label):
    lines = [label + ' {{%s.name_short}}' % role, 'Юридический адрес: {{%s.addr_legal}}' % role, 'Фактический адрес: {{%s.addr_fact}}' % role,
             '{{%s.reg_line}}' % role, 'Р/с {{%s.rs}}' % role, 'К/с {{%s.ks}}' % role, '{{%s.bank_name}}' % role, 'БИК {{%s.bik}}' % role, 'Эл. почта: {{%s.email}}' % role]
    for i, p in enumerate(ps):
        if i < len(lines): settext(p, lines[i])
        else: drop(p)
    # первая строка — жирная целиком
    for r in ps[0].runs: r.bold = True
block(req_exec, 'executor', 'Подрядчик'); block(req_cust, 'customer', 'Заказчик')
settext(sig_exec, 'Подрядчик ______________ {{executor.fio_short}}'); settext(sig_cust, 'Заказчик ______________ {{customer.fio_short}}')
d.save(out)
print('ok')
