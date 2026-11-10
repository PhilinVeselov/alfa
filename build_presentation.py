from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR


OUT = "Alfa_Ryadom_Final_Draft.pptx"

W = 13.333
H = 7.5

RED = RGBColor(239, 49, 36)
DARK = RGBColor(24, 24, 27)
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(246, 246, 247)
MID = RGBColor(103, 103, 112)
LINE = RGBColor(221, 221, 225)
GREEN = RGBColor(40, 167, 69)
YELLOW = RGBColor(245, 166, 35)
ORANGE = RGBColor(239, 108, 0)
BLUE = RGBColor(47, 105, 200)


prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
blank = prs.slide_layouts[6]


def rect(slide, x, y, w, h, fill, radius=False, line=None):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = line if line else fill
    if radius:
        try:
            s.adjustments[0] = 0.12
        except Exception:
            pass
    return s


def textbox(slide, x, y, w, h, text, size=18, color=DARK, bold=False,
            font="Arial", align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
            margin=0.03, line_spacing=1.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def bullets(slide, x, y, w, h, items, size=17, color=DARK, bullet_color=None,
            gap=7, leading=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.name = "Arial"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.level = 0
        p.space_after = Pt(gap)
        p.line_spacing = leading or 1.0
        p.text = "•  " + item
    return box


def header(slide, number, title, kicker=None):
    rect(slide, 0, 0, W, H, WHITE)
    rect(slide, 0, 0, 0.16, H, RED)
    if kicker:
        textbox(slide, 0.55, 0.32, 7.8, 0.27, kicker.upper(), 10, RED, True)
        title_y = 0.62
    else:
        title_y = 0.37
    textbox(slide, 0.55, title_y, 11.8, 0.75, title, 28, DARK, True)
    textbox(slide, 12.35, 0.36, 0.42, 0.28, f"{number:02d}", 11, MID, True, align=PP_ALIGN.RIGHT)
    rect(slide, 0.55, 1.33, 12.15, 0.015, LINE)


def footer(slide, source=None):
    if source:
        textbox(slide, 0.56, 7.10, 10.8, 0.19, source, 8, MID)
    textbox(slide, 11.55, 7.10, 1.13, 0.19, "АЛЬФА РЯДОМ", 8, RED, True, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body=None, accent=RED, value=None,
         title_size=14, body_size=12):
    rect(slide, x, y, w, h, LIGHT, True, LINE)
    rect(slide, x, y, 0.06, h, accent, True, accent)
    if value:
        textbox(slide, x + 0.24, y + 0.18, w - 0.42, 0.48, value, 25, accent, True)
        textbox(slide, x + 0.24, y + 0.72, w - 0.42, 0.40, title, title_size, DARK, True)
        if body:
            textbox(slide, x + 0.24, y + 1.16, w - 0.42, h - 1.30, body, body_size, MID)
    else:
        textbox(slide, x + 0.24, y + 0.18, w - 0.42, 0.40, title, title_size, DARK, True)
        if body:
            textbox(slide, x + 0.24, y + 0.68, w - 0.42, h - 0.82, body, body_size, MID)


def pill(slide, x, y, w, text, fill=LIGHT, color=DARK):
    rect(slide, x, y, w, 0.34, fill, True, fill)
    textbox(slide, x + 0.04, y + 0.06, w - 0.08, 0.18, text, 9, color, True, align=PP_ALIGN.CENTER)


def slide1():
    s = prs.slides.add_slide(blank)
    rect(s, 0, 0, W, H, DARK)
    rect(s, 0, 0, 0.22, H, RED)
    pill(s, 0.75, 0.66, 2.18, "КОНКУРСНОЕ РЕШЕНИЕ", RED, WHITE)
    textbox(s, 0.75, 1.25, 8.0, 0.95, "АЛЬФА РЯДОМ", 44, WHITE, True)
    textbox(s, 0.75, 2.20, 9.1, 0.85, "10 млн карт за год — без 10 млн\nиндивидуальных доставок", 27, WHITE, True)
    textbox(s, 0.75, 3.30, 8.4, 0.60, "Распределённый запас неименного пластика и управляемая партнёрская выдача по всей России", 16, RGBColor(205,205,210))
    vals = [("10,0 млн", "успешных выдач"), ("5,5 млн", "альтернативно"), ("2–3 дня", "до получения"), ("70%", "в М12")]
    for i,(v,l) in enumerate(vals):
        x=0.75+i*2.82
        textbox(s,x,5.15,2.45,0.42,v,24,RED,True)
        textbox(s,x,5.65,2.45,0.30,l,12,WHITE,True)
    textbox(s,0.75,6.78,5.0,0.22,"Рабочая версия • 10 ноября 2026",9,RGBColor(150,150,158))


def slide2():
    s=prs.slides.add_slide(blank); header(s,2,"Масштабировать выдачу, а не штат","задача")
    card(s,0.65,1.72,3.0,1.55,"Годовая цель", "Успешно переданные физические карты", value="10,0 млн")
    card(s,3.86,1.72,3.0,1.55,"Привычный канал", "Расчётный базовый объём решения", value="4,5 млн", accent=DARK)
    card(s,7.07,1.72,3.0,1.55,"Новая инфраструктура", "Минимум по кейсу — 5,0 млн", value="5,5 млн")
    card(s,10.28,1.72,2.38,1.55,"Срок", "Преимущественно по всей стране", value="2–3 дня", accent=BLUE)
    bullets(s,0.78,3.75,5.8,2.2,["Без открытия новых отделений","Без массового найма сотрудников и собственной курьерской команды","Физический пластик остаётся частью решения"],16)
    bullets(s,6.78,3.75,5.75,2.2,["Крупные города, малые населённые пункты и удалённые территории","Идентификация, персональные данные и безопасность — обязательные ограничения","К концу года — до 70% месячного потока альтернативно"],16)
    footer(s,"Источник: конкурсный кейс, стр. 7–10 и 16–17")


def slide3():
    s=prs.slides.add_slide(blank); header(s,3,"Рынок использует части решения — мы соединяем их в систему","анализ практик")
    data=[("Офис","контроль сложных случаев","ограниченная география","резерв"),("Представитель","полная процедура","дорогой ручной масштаб","сложные сценарии"),("Торговая точка","масштаб и поток","нужна новая процедура","основа"),("ПВЗ","география и привычка","нет банковских полномочий","массовый резерв"),("Инфраоператор","малые территории","выше стоимость","география"),("Картомат","быстрая выдача","CAPEX и сроки","пилот")]
    y=1.65
    textbox(s,0.75,y,2.1,.3,"МОДЕЛЬ",10,MID,True); textbox(s,3.0,y,2.7,.3,"СИЛЬНАЯ СТОРОНА",10,MID,True); textbox(s,6.0,y,3.2,.3,"ОГРАНИЧЕНИЕ",10,MID,True); textbox(s,9.65,y,2.3,.3,"РОЛЬ",10,MID,True)
    for i,row in enumerate(data):
        yy=2.05+i*.72
        rect(s,.68,yy,11.95,.58,WHITE if i%2==0 else LIGHT,True,LINE)
        textbox(s,.82,yy+.13,1.95,.25,row[0],13,DARK,True)
        textbox(s,3.0,yy+.13,2.72,.25,row[1],12,DARK)
        textbox(s,6.0,yy+.13,3.25,.25,row[2],12,MID)
        pill(s,9.65,yy+.11,2.25,row[3], RED if row[3]=="основа" else LIGHT, WHITE if row[3]=="основа" else DARK)
    footer(s,"Вывод: один канал не закрывает одновременно объём, географию, безопасность и срок")


def slide4():
    s=prs.slides.add_slide(blank); header(s,4,"Переместить карту заранее — не везти после каждого заказа","главный инсайт")
    steps=[("1","Производство"),("2","Федеральный\nрезерв"),("3","Региональный\nузел"),("4","Запас\nв точке"),("5","Заявка и\nрезерв"),("6","Разрешение\nбанка"),("7","Передача\nклиенту")]
    for i,(n,t) in enumerate(steps):
        x=.55+i*1.79
        rect(s,x,2.08,1.46,1.32,LIGHT,True,LINE)
        rect(s,x+.12,2.23,.38,.38,RED,True,RED)
        textbox(s,x+.12,2.31,.38,.16,n,10,WHITE,True,align=PP_ALIGN.CENTER)
        textbox(s,x+.12,2.72,1.18,.47,t,12,DARK,True,align=PP_ALIGN.CENTER)
        if i<6:
            textbox(s,x+1.48,2.58,.26,.24,"→",18,RED,True,align=PP_ALIGN.CENTER)
    card(s,.76,4.02,3.75,1.55,"До заказа", "Неименной пластик уже находится в регионе или точке; запас пополняется по прогнозу.",accent=BLUE)
    card(s,4.78,4.02,3.75,1.55,"После заказа", "Система резервирует конкретную карту и показывает только реально доступный маршрут.",accent=RED)
    card(s,8.80,4.02,3.75,1.55,"До передачи", "Карта не связана со счётом и не предоставляет доступ к средствам клиента.",accent=GREEN)
    textbox(s,.78,6.02,11.7,.55,"Так SLA 2–3 дня достигается даже там, где индивидуальная доставка из центра заняла бы дольше.",19,DARK,True,align=PP_ALIGN.CENTER)
    footer(s)


def slide5():
    s=prs.slides.add_slide(blank); header(s,5,"Единая платформа показывает только выполнимое обещание","целевая модель")
    rect(s,4.47,2.43,4.42,1.45,DARK,True,DARK)
    textbox(s,4.77,2.72,3.82,.40,"АЛЬФА РЯДОМ",24,WHITE,True,align=PP_ALIGN.CENTER)
    textbox(s,4.77,3.22,3.82,.25,"операционная платформа",11,RGBColor(200,200,205),align=PP_ALIGN.CENTER)
    cards=[(.65,1.65,"Реестр","Карты, остатки, статусы",BLUE),(9.13,1.65,"Прогноз","Спрос и пополнение",GREEN),(.65,4.15,"Маршрутизация","Срок, расстояние, мощность",RED),(9.13,4.15,"Контроль","Выдача, возвраты, инциденты",ORANGE),(4.47,4.55,"Продвижение","Трафик только в свободную мощность",RED)]
    for x,y,t,b,a in cards: card(s,x,y,3.55,1.26,t,b,accent=a,title_size=14,body_size=11)
    for x1,y1,x2,y2 in [(4.2,2.25,4.47,2.75),(8.89,2.75,9.13,2.25),(4.2,4.55,4.47,3.55),(8.89,3.55,9.13,4.55)]:
        textbox(s,min(x1,x2),min(y1,y2),abs(x2-x1)+.2,abs(y2-y1)+.2,"↔",18,MID,True,align=PP_ALIGN.CENTER)
    footer(s,"Критерий показа точки: запас + мощность + допустимый процесс + срок + экономика")


def slide6():
    s=prs.slides.add_slide(blank); header(s,6,"Вся Россия — но не один одинаковый канал","география")
    rows=[("АГЛОМЕРАЦИИ","X5 • ПВЗ • картомат","доставка • офис",RED),("СРЕДНИЕ И МАЛЫЕ ГОРОДА","X5 • ПВЗ","инфраструктурный оператор",BLUE),("ПОСЁЛКИ И СЁЛА","партнёрская точка • оператор","групповой маршрут",GREEN),("ТРУДНОДОСТУПНЫЕ ТЕРРИТОРИИ","предварительный запас у оператора","доставка по расписанию",ORANGE)]
    for i,(seg,main,res,a) in enumerate(rows):
        y=1.66+i*1.18
        rect(s,.68,y,12.0,.95,LIGHT,True,LINE)
        rect(s,.68,y,.14,.95,a,True,a)
        textbox(s,1.00,y+.18,3.15,.23,seg,12,DARK,True)
        textbox(s,4.35,y+.14,4.15,.22,"ОСНОВНОЙ МАРШРУТ",9,MID,True)
        textbox(s,4.35,y+.43,4.15,.25,main,14,DARK,True)
        textbox(s,8.95,y+.14,3.05,.22,"РЕЗЕРВ",9,MID,True)
        textbox(s,8.95,y+.43,3.05,.25,res,13,DARK)
    textbox(s,.78,6.46,11.75,.42,"В удалённых территориях срок обеспечивает заранее размещённый запас на 21–30 дней, а не срочная перевозка после заявки.",16,DARK,True,align=PP_ALIGN.CENTER)
    footer(s,"X5: 29 790 магазинов в 75 субъектах РФ на 31.12.2025 — охват значительный, но не полный")


def slide7():
    s=prs.slides.add_slide(blank); header(s,7,"5,5 млн альтернативных выдач — с запасом над минимумом","структура объёма")
    items=[("X5 и торговые точки",3.0,RED),("ПВЗ",1.2,BLUE),("Инфраструктурный оператор",.7,GREEN),("Корпоративная выдача",.4,ORANGE),("Картоматы",.2,MID),("Традиционная выдача",4.5,DARK)]
    maxv=4.5
    for i,(name,val,col) in enumerate(items):
        y=1.63+i*.72
        textbox(s,.72,y+.08,2.85,.26,name,12,DARK,True if i in (0,5) else False)
        rect(s,3.65,y,7.4,.44,LIGHT,True,LIGHT)
        rect(s,3.65,y,7.4*val/maxv,.44,col,True,col)
        textbox(s,11.28,y+.08,1.0,.23,f"{val:.1f} млн".replace('.',','),13,col,True,align=PP_ALIGN.RIGHT)
    card(s,9.27,5.98,3.08,.68,"Резерв над минимумом",value="+0,5 млн",accent=RED,title_size=11)
    textbox(s,.75,6.15,7.7,.36,"Все объёмы каналов — расчётные гипотезы до подтверждения пилотом и договором.",12,MID)
    footer(s)


def slide8():
    s=prs.slides.add_slide(blank); header(s,8,"План достижим при резерве мощности 1,231 млн","capacity planning")
    vals=[("6,731 млн","расчётная альтернативная мощность",BLUE),("5,500 млн","план альтернативных выдач",RED),("1,231 млн","операционный резерв",GREEN),("≤85,5%","максимальная загрузка канала",ORANGE)]
    for i,(v,l,c) in enumerate(vals): card(s,.66+i*3.05,1.66,2.77,1.48,l,value=v,accent=c,title_size=12)
    pointdata=[("X5","14 000"),("ПВЗ","6 500"),("Инфраточки","2 300"),("Корп. точки","105"),("Картоматы","200")]
    for i,(n,v) in enumerate(pointdata):
        x=.74+i*2.42
        rect(s,x,3.73,2.10,1.28,LIGHT,True,LINE)
        textbox(s,x+.15,3.96,1.80,.35,v,23,DARK,True,align=PP_ALIGN.CENTER)
        textbox(s,x+.15,4.47,1.80,.25,n,11,MID,True,align=PP_ALIGN.CENTER)
    textbox(s,.82,5.55,11.65,.48,"Успешные выдачи = активные точки × выдачи/день × рабочие дни × доступность × конверсия",16,DARK,True,align=PP_ALIGN.CENTER)
    pill(s,4.55,6.27,4.25,"КАЖДОЕ ДОПУЩЕНИЕ ПРОВЕРЯЕТСЯ ПИЛОТОМ",RED,WHITE)
    footer(s)


def slide9():
    s=prs.slides.add_slide(blank); header(s,9,"Путь клиента заканчивается первой транзакцией — не доставкой в точку","CJM")
    stages=["Предложение","Заявка","Статус","Выбор","Резерв","Готовность","Проверка","Передача","Активация","1-я операция"]
    for i,t in enumerate(stages):
        x=.55+i*1.25
        c=RED if i in (0,4,7,9) else DARK
        rect(s,x,1.72,1.02,.56,c,True,c)
        textbox(s,x+.03,1.88,.96,.18,t,9,WHITE,True,align=PP_ALIGN.CENTER)
        if i<9: textbox(s,x+1.04,1.86,.18,.20,"→",13,RED,True,align=PP_ALIGN.CENTER)
    layers=[("КЛИЕНТ","выбирает удобный реальный маршрут • получает понятный срок • активирует карту",RED), ("БАНК","определяет допустимость • резервирует • разрешает передачу • поддерживает",BLUE),("ПАРТНЁР","хранит защищённый запас • выполняет только разрешённую операцию • фиксирует передачу",GREEN)]
    for i,(t,b,c) in enumerate(layers):
        y=2.86+i*.91
        rect(s,.72,y,11.9,.68,LIGHT,True,LINE); rect(s,.72,y,.12,.68,c,True,c)
        textbox(s,1.02,y+.18,1.45,.20,t,10,c,True)
        textbox(s,2.52,y+.16,9.55,.24,b,13,DARK)
    card(s,.72,5.86,3.75,.76,"Затруднение","Нет карты • очередь • ошибка • сбой",accent=ORANGE,title_size=11,body_size=10)
    card(s,4.78,5.86,3.75,.76,"Помощь","Соседняя точка • чат • видео • другой канал",accent=GREEN,title_size=11,body_size=10)
    card(s,8.84,5.86,3.75,.76,"Контроль","Конверсия • SLA • ожидание • активация",accent=BLUE,title_size=11,body_size=10)
    footer(s)


def slide10():
    s=prs.slides.add_slide(blank); header(s,10,"Начинаем с безопасного сегмента — расширяем через Stage-Gate","идентификация")
    card(s,.72,1.66,5.65,3.18,"A. Действующий идентифицированный клиент","1. Банк подтверждает статус в собственном контуре.\n2. Клиент выбирает доступную точку.\n3. Банк разрешает передачу конкретной карты.\n4. Партнёр выполняет только утверждённые действия.\n5. Клиент активирует карту в банке.",accent=GREEN,title_size=17,body_size=14)
    card(s,6.77,1.66,5.65,3.18,"B. Новый или сложный клиент","1. Требуется предусмотренная банковская процедура.\n2. Альтернативный путь показывается только после юридического утверждения.\n3. Иначе — уполномоченный представитель, офис или другой допустимый канал.\n4. Никаких упрощений ради объёма.",accent=ORANGE,title_size=17,body_size=14)
    rect(s,.72,5.32,11.70,.84,DARK,True,DARK)
    textbox(s,1.02,5.53,11.10,.40,"Обычный сотрудник магазина или ПВЗ не получает банковские полномочия автоматически.",18,WHITE,True,align=PP_ALIGN.CENTER)
    footer(s,"Юридическая модель является обязательным G0-критерием до пилота")


def slide11():
    s=prs.slides.add_slide(blank); header(s,11,"Карту невозможно выдать без онлайн-разрешения банка","логистика и безопасность")
    sts=[("PRODUCED",BLUE),("FEDERAL",BLUE),("REGIONAL",BLUE),("POINT",GREEN),("RESERVED",ORANGE),("AUTHORIZED",RED),("HANDED OVER",RED),("ACTIVATED",DARK)]
    for i,(t,c) in enumerate(sts):
        x=.43+i*1.58
        rect(s,x,1.75,1.30,.54,c,True,c)
        textbox(s,x+.03,1.92,1.24,.17,t,8,WHITE,True,align=PP_ALIGN.CENTER)
        if i<7: textbox(s,x+1.31,1.91,.24,.18,"→",12,MID,True,align=PP_ALIGN.CENTER)
    rect(s,4.95,2.62,3.40,.58,LIGHT,True,LINE); textbox(s,5.07,2.80,3.16,.18,"QUARANTINE  →  DESTROYED",10,ORANGE,True,align=PP_ALIGN.CENTER)
    controls=[("Неактивный пластик","До передачи нет доступа к средствам"),("Индивидуальный учёт","Каждая карта и смена статуса журналируются"),("Закрытое хранение","Пломбы, модуль, личные учётные записи"),("Сверка и блокировка","Любое расхождение — карантин и расследование")]
    for i,(t,b) in enumerate(controls): card(s,.72+(i%2)*6.05,3.66+(i//2)*1.28,5.65,1.02,t,b,accent=[BLUE,GREEN,ORANGE,RED][i],title_size=13,body_size=11)
    footer(s)


def slide12():
    s=prs.slides.add_slide(blank); header(s,12,"679 ₽ за альтернативную выдачу — при базовых допущениях","экономика")
    data=[("X5",620,RED),("ПВЗ",660,BLUE),("Инфраоператор",890,GREEN),("Корпоративная",630,ORANGE),("Картомат",1030,MID)]
    maxv=1100
    for i,(n,v,c) in enumerate(data):
        y=1.64+i*.71
        textbox(s,.75,y+.10,1.72,.25,n,12,DARK,True)
        rect(s,2.55,y,6.65,.46,LIGHT,True,LIGHT); rect(s,2.55,y,6.65*v/maxv,.46,c,True,c)
        textbox(s,9.42,y+.09,.90,.24,f"{v} ₽",13,c,True,align=PP_ALIGN.RIGHT)
    card(s,10.58,1.68,2.02,1.46,"Среднее",value="679 ₽",accent=RED,title_size=12)
    card(s,10.58,3.34,2.02,1.46,"Экономия",body="против 1 000 ₽\nза 5,5 млн",value="1,767 млрд ₽",accent=GREEN,title_size=11,body_size=10)
    textbox(s,.78,5.58,8.90,.30,"Экономия создаётся партиями и существующей инфраструктурой — не покупкой устройств.",15,DARK,True)
    textbox(s,.78,6.05,11.6,.38,"* Тарифы являются учебными допущениями и заменяются фактическими значениями по результатам коммерческих предложений и пилота.",10,MID)
    footer(s,"Ориентиры кейса: 1 000 ₽ обычная доставка; до 2 000 ₽ при расширении команды")


def slide13():
    s=prs.slides.add_slide(blank); header(s,13,"Продвижение включается только при свободной мощности","управление спросом")
    modes=[("ЗЕЛЁНЫЙ","<75%","усилить показ • разрешить стимул",GREEN),("ЖЁЛТЫЙ","75–85%","стандартное продвижение",YELLOW),("ОРАНЖЕВЫЙ","85–95%","остановить платный трафик • перенаправить",ORANGE),("КРАСНЫЙ",">95%","скрыть точку для новых заказов",RED)]
    for i,(n,v,a,c) in enumerate(modes):
        y=1.64+i*1.04
        rect(s,.72,y,11.75,.80,LIGHT,True,LINE); rect(s,.72,y,.19,.80,c,True,c)
        textbox(s,1.08,y+.22,1.40,.22,n,11,c,True)
        textbox(s,2.78,y+.17,1.15,.30,v,20,DARK,True)
        textbox(s,4.25,y+.22,7.65,.23,a,14,DARK)
    textbox(s,.76,6.08,11.7,.37,"Лестница ценности: рядом → быстрее → гибко → подтверждённое наличие → бонус только при необходимости",16,DARK,True,align=PP_ALIGN.CENTER)
    footer(s,"Нельзя выполнять маркетинговый KPI ценой очереди, stockout или нарушения SLA")


def slide14():
    s=prs.slides.add_slide(blank); header(s,14,"Альтернативный поток растёт до 70% в М12","план масштабирования")
    total=[.25,.35,.50,.65,.75,.85,.90,1.00,1.05,1.15,1.25,1.30]
    alt=[.01,.03,.09,.20,.34,.47,.58,.65,.69,.74,.79,.91]
    maxv=1.30
    base_y=5.75; chart_h=3.75
    for i,(t,a) in enumerate(zip(total,alt)):
        x=.72+i*.98
        bh=chart_h*t/maxv; ah=chart_h*a/maxv
        rect(s,x,base_y-bh,.56,bh,DARK,True,DARK)
        rect(s,x,base_y-ah,.56,ah,RED,True,RED)
        textbox(s,x-.02,5.92,.60,.22,f"М{i+1}",9,MID,True,align=PP_ALIGN.CENTER)
    textbox(s,.78,1.57,1.55,.22,"млн выдач",9,MID,True)
    pill(s,9.88,1.50,1.06,"АЛЬТ",RED,WHITE); pill(s,11.04,1.50,1.06,"ВСЕГО",DARK,WHITE)
    checkpoints=[("М2","закрытый пилот"),("М4","0,20 млн альтернативно"),("М6","55% в месяце"),("М9","forecast ≥9,7 млн"),("М12","0,91 млн / 70%")]
    for i,(m,t) in enumerate(checkpoints):
        x=.72+i*2.40
        textbox(s,x,6.49,2.18,.18,m,9,RED,True,align=PP_ALIGN.CENTER)
        textbox(s,x,6.74,2.18,.22,t,9,DARK,True,align=PP_ALIGN.CENTER)
    footer(s,"Итого за год: 10,0 млн • 5,5 млн альтернативно • 4,5 млн традиционно")


def slide15():
    s=prs.slides.add_slide(blank); header(s,15,"Масштабирование открывают доказательства, а не календарь","Stage-Gate")
    gates=[("G0","Дизайн","право • данные • партнёры"),("G1","Готовность","end-to-end • запас • безопасность"),("G2","Пилот","спрос • SLA • стоимость"),("G3","Тираж","470 тыс./мес • резервный оператор"),("G4","Федерально","forecast • запас • риски"),("G5","Результат","10 млн • аудит • следующий год")]
    for i,(g,t,b) in enumerate(gates):
        x=.60+i*2.10
        c=RED if i in (0,2,5) else DARK
        rect(s,x,2.02,1.72,2.55,LIGHT,True,LINE)
        rect(s,x+.52,1.65,.68,.68,c,True,c)
        textbox(s,x+.52,1.88,.68,.18,g,11,WHITE,True,align=PP_ALIGN.CENTER)
        textbox(s,x+.17,2.40,1.38,.32,t,15,DARK,True,align=PP_ALIGN.CENTER)
        textbox(s,x+.17,3.05,1.38,.92,b,11,MID,align=PP_ALIGN.CENTER)
        if i<5: textbox(s,x+1.75,3.00,.32,.28,"→",18,RED,True,align=PP_ALIGN.CENTER)
    rect(s,.73,5.34,11.72,.80,DARK,True,DARK)
    textbox(s,1.02,5.54,11.12,.38,"Открытый критический риск = следующий масштаб не запускается.",18,WHITE,True,align=PP_ALIGN.CENTER)
    footer(s)


def slide16():
    s=prs.slides.add_slide(blank); header(s,16,"Управляем успешной выдачей — не выпуском пластика","KPI")
    groups=[("РЕЗУЛЬТАТ",["успешно выдано","альтернативная доля","активация","первая транзакция"],RED),("ВОРОНКА",["заявка → резерв","резерв → получение","неявка","повторный визит"],BLUE),("ОПЕРАЦИИ",["SLA 2–3 дня","мощность и загрузка","точность остатка","stockout"],GREEN),("РИСК И ЭКОНОМИКА",["стоимость выдачи","ошибки и утраты","жалобы","rolling forecast"],ORANGE)]
    for i,(t,items,c) in enumerate(groups):
        x=.72+(i%2)*6.05; y=1.65+(i//2)*2.33
        rect(s,x,y,5.65,1.92,LIGHT,True,LINE); rect(s,x,y,.10,1.92,c,True,c)
        textbox(s,x+.28,y+.20,5.05,.25,t,11,c,True)
        bullets(s,x+.28,y+.62,5.02,1.08,items,12,DARK,gap=3)
    rect(s,.72,6.05,11.70,.55,RED,True,RED)
    textbox(s,1.00,6.20,11.14,.20,"«Выпущено» и «доставлено в точку» не заменяют KPI «передано клиенту».",15,WHITE,True,align=PP_ALIGN.CENTER)
    footer(s)


def slide17():
    s=prs.slides.add_slide(blank); header(s,17,"Риски получают владельца и ранний сигнал до масштабирования","risk management")
    risks=[("80% выбирают курьера","Продукт + маркетинг",9,RED),("Юрмодель не согласована","Юристы / комплаенс",6,ORANGE),("Партнёр или интеграция задержаны","Партнёрства + ИТ",6,ORANGE),("Мощность ниже расчёта","Операционный центр",6,ORANGE),("Stockout при резерве","Логистика",6,ORANGE),("Мошенничество / ошибочная выдача","Антифрод",6,ORANGE),("Концентрация на X5","Руководитель программы",6,ORANGE),("Часть географии вне SLA","Региональные операции",6,ORANGE)]
    for i,(r,o,score,c) in enumerate(risks):
        x=.70+(i%2)*6.08; y=1.61+(i//2)*1.18
        rect(s,x,y,5.67,.94,LIGHT,True,LINE)
        rect(s,x+.16,y+.18,.55,.55,c,True,c)
        textbox(s,x+.16,y+.36,.55,.17,str(score),11,WHITE,True,align=PP_ALIGN.CENTER)
        textbox(s,x+.88,y+.17,4.52,.25,r,13,DARK,True)
        textbox(s,x+.88,y+.52,4.52,.20,o,10,MID)
    textbox(s,.76,6.48,11.62,.28,"Право, безопасность, утрата и ошибочная выдача требуют реакции независимо от математического балла.",12,DARK,True,align=PP_ALIGN.CENTER)
    footer(s)


def slide18():
    s=prs.slides.add_slide(blank); header(s,18,"Два кризиса — два заранее подготовленных ответа","устойчивость")
    card(s,.70,1.62,5.75,4.42,"1. 80% выбирают курьера","Сигнал: курьер >70% семь дней; загрузка >90%.\n\n0–48 ч: проверить доступность, интерфейс и доверие.\nНеделя 1: перенаправить CRM в свободную мощность.\nНеделя 2: тест скорости и порядка показа.\nДалее: бонус только для нужных сегментов.\n\nВосстановление: альтернатива ≥50%, курьер <85% две недели.",accent=RED,title_size=18,body_size=14)
    card(s,6.83,1.62,5.75,4.42,"2. Forecast падает с 10 до 8 млн","Компенсация разрыва 2,0 млн:\n\n+0,500 млн — традиционный резерв до лимита 5 млн;\n+0,511 млн — резерв работающих альтернатив;\n+1,010 млн — заранее подготовленные 5 тыс. точек.\n\nВосстановление: forecast 10 млн четыре недели, альтернативы ≥5 млн.",accent=ORANGE,title_size=18,body_size=14)
    textbox(s,.85,6.43,11.55,.35,"АЛЬФА РЯДОМ масштабирует доступность карты — не индивидуальную доставку.",19,DARK,True,align=PP_ALIGN.CENTER)
    footer(s)


def appendix_header(slide, number, title):
    header(slide, number, title, "приложение")


def slide19():
    s=prs.slides.add_slide(blank); appendix_header(s,19,"Допущения модели мощности")
    rows=[("X5","2,0","30","97%","82%","47,7"),("ПВЗ","1,5","30","96%","78%","33,7"),("Инфраоператор","3,0","26","95%","85%","63,0"),("Корпоративная","20,0","20","95%","90%","342,0"),("Картомат","10,0","30","95%","85%","242,3")]
    headers=[("КАНАЛ",.75,2.25),("В ДЕНЬ",3.16,1.15),("ДНЕЙ",4.45,1.05),("ДОСТУПНОСТЬ",5.68,1.48),("КОНВЕРСИЯ",7.37,1.38),("В МЕСЯЦ",9.08,1.50)]
    for t,x,w in headers: textbox(s,x,1.62,w,.22,t,9,MID,True,align=PP_ALIGN.CENTER)
    for i,row in enumerate(rows):
        y=2.04+i*.72; rect(s,.70,y,10.0,.56,WHITE if i%2==0 else LIGHT,True,LINE)
        vals=[(row[0],.84,2.10,PP_ALIGN.LEFT),(row[1],3.16,1.15,PP_ALIGN.CENTER),(row[2],4.45,1.05,PP_ALIGN.CENTER),(row[3],5.68,1.48,PP_ALIGN.CENTER),(row[4],7.37,1.38,PP_ALIGN.CENTER),(row[5],9.08,1.50,PP_ALIGN.CENTER)]
        for v,x,w,a in vals: textbox(s,x,y+.16,w,.20,v,12,DARK,True if x==.84 else False,align=a)
    card(s,10.94,1.65,1.70,2.08,"Формула","точки × выдачи/день × дни × доступность × конверсия",accent=RED,title_size=12,body_size=11)
    card(s,10.94,3.97,1.70,1.64,"Проверка","база • −15% • −20% • задержка канала",accent=BLUE,title_size=12,body_size=10)
    textbox(s,.78,6.18,11.65,.43,"Все параметры производительности — учебные допущения до подтверждения четырьмя неделями стабильного пилота.",13,DARK,True,align=PP_ALIGN.CENTER)
    footer(s,"Источник расчёта: 02_Model_moshchnosti.md")


def slide20():
    s=prs.slides.add_slide(blank); appendix_header(s,20,"Помесячный план по каналам, млн выдач")
    rows=[("М1","0,25","0,01","0,24","4%"),("М2","0,35","0,03","0,32","9%"),("М3","0,50","0,09","0,41","18%"),("М4","0,65","0,20","0,45","31%"),("М5","0,75","0,34","0,41","45%"),("М6","0,85","0,47","0,38","55%"),("М7","0,90","0,58","0,32","64%"),("М8","1,00","0,65","0,35","65%"),("М9","1,05","0,69","0,36","66%"),("М10","1,15","0,74","0,41","64%"),("М11","1,25","0,79","0,46","63%"),("М12","1,30","0,91","0,39","70%"),("ИТОГО","10,00","5,50","4,50","55%")]
    headers=[("МЕСЯЦ",.72,1.25),("ВСЕГО",2.25,1.55),("АЛЬТЕРНАТИВНО",4.10,2.05),("ТРАДИЦИОННО",6.50,2.0),("ДОЛЯ АЛЬТ.",8.83,1.75)]
    for t,x,w in headers: textbox(s,x,1.52,w,.22,t,9,MID,True,align=PP_ALIGN.CENTER)
    for i,row in enumerate(rows):
        y=1.88+i*.37; is_total=i==12
        rect(s,.68,y,10.1,.30,RED if is_total else (WHITE if i%2==0 else LIGHT),True,RED if is_total else LINE)
        cols=[(.76,1.10),(2.32,1.40),(4.25,1.70),(6.64,1.70),(9.00,1.45)]
        for j,(x,w) in enumerate(cols): textbox(s,x,y+.075,w,.14,row[j],9,WHITE if is_total else DARK,True if is_total or j==0 else False,align=PP_ALIGN.CENTER)
    card(s,11.05,1.88,1.55,1.54,"М12",body="0,91 млн альтернативно",value="70%",accent=RED,title_size=11,body_size=10)
    card(s,11.05,3.66,1.55,1.54,"Год",body="резерв 0,5 млн над минимумом",value="5,5",accent=GREEN,title_size=11,body_size=9)
    footer(s,"Источник: capacity_plan.csv")


def slide21():
    s=prs.slides.add_slide(blank); appendix_header(s,21,"Критические развилки CJM")
    cases=[("Нет карты","блокировка точки → соседняя точка → компенсация SLA","stockout • перенос • срок",RED),("Не прошёл проверку","карта не передаётся → повтор или допустимый очный канал","отказы • ложные отказы",ORANGE),("Точка перегружена","скрыть новый спрос → сохранить резервы → добавить мощность","загрузка • очередь • SLA",BLUE),("Клиент не пришёл","напоминание → продление → снятие резерва","no-show • замороженный запас",GREEN),("Повреждение","карантин → замена → обратная логистика","повреждения • время замены",ORANGE),("Нет связи","выдача запрещена → уведомление → другой канал","доступность • восстановление",RED)]
    for i,(t,a,m,c) in enumerate(cases):
        x=.70+(i%2)*6.04; y=1.58+(i//2)*1.63
        rect(s,x,y,5.64,1.36,LIGHT,True,LINE); rect(s,x,y,.10,1.36,c,True,c)
        textbox(s,x+.28,y+.18,5.03,.24,t,14,DARK,True)
        textbox(s,x+.28,y+.54,5.03,.35,a,11,DARK)
        textbox(s,x+.28,y+1.04,5.03,.17,m,9,c,True)
    footer(s,"Полная CJM: 03_CJM_Alfa_Ryadom.md")


def slide22():
    s=prs.slides.add_slide(blank); appendix_header(s,22,"RACI: кто отвечает за движение карты")
    rows=[("Производство и учёт","A","R","—","—"),("Межрегиональная перевозка","A","—","R","—"),("Региональное хранение","A","—","C","R"),("Приёмка и хранение в точке","A","—","C","R"),("Банковская проверка","A/R","—","—","C"),("Физическая передача","A","—","—","R"),("Возврат и обратная логистика","A","—","R","R"),("Расследование инцидента","A/R","C","C","C")]
    headers=[("ПРОЦЕСС",.72,4.3),("БАНК",5.20,1.25),("ПРОИЗВ.",6.58,1.25),("3PL",7.96,1.25),("ТОЧКА",9.34,1.25)]
    for t,x,w in headers:textbox(s,x,1.55,w,.22,t,9,MID,True,align=PP_ALIGN.CENTER)
    for i,row in enumerate(rows):
        y=1.92+i*.54; rect(s,.68,y,10.05,.40,WHITE if i%2==0 else LIGHT,True,LINE)
        textbox(s,.82,y+.10,4.10,.18,row[0],10,DARK,True if i in (0,4,5) else False)
        for j,x in enumerate([5.20,6.58,7.96,9.34],1): textbox(s,x,y+.10,1.25,.18,row[j],10,RED if "A" in row[j] else DARK,True,align=PP_ALIGN.CENTER)
    card(s,11.00,1.78,1.60,2.35,"Обозначения","R — выполняет\nA — отвечает\nC — согласует\nI — информируется",accent=RED,title_size=11,body_size=11)
    card(s,11.00,4.36,1.60,1.46,"Принцип","банк сохраняет ответственность за банковский результат",accent=DARK,title_size=11,body_size=10)
    footer(s,"Юридическая и материальная ответственность закрепляются договорами")


def slide23():
    s=prs.slides.add_slide(blank); appendix_header(s,23,"Unit-экономика по каналам")
    rows=[("X5","3,0","360","500","620"),("ПВЗ","1,2","120","560","660"),("Инфраоператор","0,7","98","750","890"),("Корпоративная","0,4","24","570","630"),("Картомат","0,2","150","280","1 030"),("АЛЬТЕРНАТИВЫ","5,5","752","542","679")]
    headers=[("КАНАЛ",.72,2.4),("ОБЪЁМ, МЛН",3.32,1.55),("FIX, МЛН ₽",5.03,1.55),("VAR, ₽",6.77,1.35),("ИТОГО, ₽",8.35,1.50)]
    for t,x,w in headers:textbox(s,x,1.58,w,.22,t,9,MID,True,align=PP_ALIGN.CENTER)
    for i,row in enumerate(rows):
        y=1.96+i*.63; total=i==5; rect(s,.68,y,9.35,.48,RED if total else (WHITE if i%2==0 else LIGHT),True,RED if total else LINE)
        coords=[(.82,2.20),(3.35,1.50),(5.07,1.50),(6.80,1.30),(8.38,1.45)]
        for j,(x,w) in enumerate(coords): textbox(s,x,y+.14,w,.19,row[j],11,WHITE if total else DARK,True if total or j==0 else False,align=PP_ALIGN.CENTER if j else PP_ALIGN.LEFT)
    card(s,10.34,1.92,2.25,1.54,"Экономия",body="против 1 000 ₽",value="1,767 млрд ₽",accent=GREEN,title_size=11,body_size=10)
    card(s,10.34,3.69,2.25,1.54,"Программа",body="10 млн выдач",value="10,733 млрд ₽",accent=RED,title_size=11,body_size=10)
    textbox(s,.78,6.10,11.6,.30,"Картомат остаётся пилотом: покупка оборудования повышает стрессовую стоимость примерно до 2 380 ₽.",12,DARK,True,align=PP_ALIGN.CENTER)
    footer(s,"Все тарифы — учебные допущения. Источник: unit_economics.csv")


def slide24():
    s=prs.slides.add_slide(blank); appendix_header(s,24,"Анализ чувствительности экономики")
    scenarios=[("БАЗА","5,5 млн","679 ₽","3,733 млрд ₽",GREEN),("ОБЪЁМ −20%","4,4 млн","713 ₽","3,137 млрд ₽",YELLOW),("VAR +20%","5,5 млн","787 ₽","4,329 млрд ₽",ORANGE),("ОБЪЁМ −20% + VAR +20%","4,4 млн","821 ₽","3,614 млрд ₽",RED)]
    for i,(n,v,cost,total,c) in enumerate(scenarios):
        y=1.60+i*1.16; rect(s,.72,y,11.75,.90,LIGHT,True,LINE); rect(s,.72,y,.16,.90,c,True,c)
        textbox(s,1.08,y+.16,3.15,.20,n,12,DARK,True)
        textbox(s,4.52,y+.14,1.62,.30,v,18,DARK,True,align=PP_ALIGN.CENTER)
        textbox(s,6.52,y+.14,1.62,.30,cost,18,c,True,align=PP_ALIGN.CENTER)
        textbox(s,8.58,y+.14,2.45,.30,total,18,DARK,True,align=PP_ALIGN.CENTER)
        textbox(s,4.52,y+.52,1.62,.15,"объём",8,MID,align=PP_ALIGN.CENTER); textbox(s,6.52,y+.52,1.62,.15,"за выдачу",8,MID,align=PP_ALIGN.CENTER); textbox(s,8.58,y+.52,2.45,.15,"затраты",8,MID,align=PP_ALIGN.CENTER)
    textbox(s,.78,6.42,11.6,.28,"Экономика выдерживает стресс, но снижение объёма до 4,4 млн нарушает обязательное условие кейса.",12,DARK,True,align=PP_ALIGN.CENTER)
    footer(s)


def slide25():
    s=prs.slides.add_slide(blank); appendix_header(s,25,"Триггеры автоматического управления")
    items=[("Forecast <9,5 млн","подготовить резервный канал",YELLOW),("Forecast <9,0 млн","частично активировать резерв",ORANGE),("Forecast =8,0 млн","полный антикризисный режим",RED),("Загрузка >85%","остановить дополнительное продвижение",ORANGE),("Загрузка >95%","скрыть точку для новых заказов",RED),("Stockout >0,5%","сверка и ограничение резервов",ORANGE),("Стоимость >1 000 ₽","пересмотреть договор или остановить расширение",YELLOW),("Ошибка / утрата","немедленно остановить точку или партию",RED)]
    for i,(tr,act,c) in enumerate(items):
        x=.72+(i%2)*6.04; y=1.58+(i//2)*1.18
        rect(s,x,y,5.64,.92,LIGHT,True,LINE); rect(s,x,y,.10,.92,c,True,c)
        textbox(s,x+.28,y+.16,1.82,.22,tr,12,c,True)
        textbox(s,x+2.18,y+.16,3.05,.42,act,11,DARK)
    footer(s,"Источник: 08_Roadmap_i_KPI.md")


def slide26():
    s=prs.slides.add_slide(blank); appendix_header(s,26,"Источники, факты и границы утверждений")
    cols=[("ДАНО КЕЙСОМ",["10 млн / 12 месяцев","≥5 млн альтернативно","2–3 дня","1 000–2 000 ₽","два кризисных сценария"],RED),("ОФИЦИАЛЬНЫЕ ИСТОЧНИКИ",["X5: география и число точек","Альфа-Банк: текущие способы","Ozon: масштаб ПВЗ","Сбер / Т-Банк: рыночные практики"],BLUE),("УЧЕБНЫЕ ДОПУЩЕНИЯ",["объёмы по каналам","производительность точек","конверсия","тарифы и экономика","резервный пул"],ORANGE)]
    for i,(t,items,c) in enumerate(cols):
        x=.68+i*4.18; rect(s,x,1.62,3.82,3.42,LIGHT,True,LINE); rect(s,x,1.62,.10,3.42,c,True,c)
        textbox(s,x+.28,1.88,3.24,.26,t,11,c,True)
        bullets(s,x+.28,2.40,3.18,2.22,items,11,DARK,gap=7)
    rect(s,.70,5.48,11.75,.80,DARK,True,DARK)
    textbox(s,.96,5.64,11.20,.48,"Нельзя утверждать согласие партнёра, юридическую допустимость или фактический тариф без договора, заключения или пилота.",14,WHITE,True,align=PP_ALIGN.CENTER)
    footer(s,"Полный реестр: 10_Reestr_istochnikov_i_dopushcheniy.md")


for fn in [slide1,slide2,slide3,slide4,slide5,slide6,slide7,slide8,slide9,slide10,slide11,slide12,slide13,slide14,slide15,slide16,slide17,slide18,slide19,slide20,slide21,slide22,slide23,slide24,slide25,slide26]:
    fn()

# Core properties
prs.core_properties.title = "Альфа Рядом — конкурсное решение"
prs.core_properties.subject = "Модель выдачи 10 млн дебетовых карт за 12 месяцев"
prs.core_properties.author = "Проектная команда"
prs.core_properties.keywords = "Альфа-Банк, выдача карт, партнёрская сеть, логистика"
prs.save(OUT)
print(f"saved {OUT}: {len(prs.slides)} slides")
