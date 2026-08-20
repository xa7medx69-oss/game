from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from build_game_explainer_ppt import (
    AMBER,
    ASSET_DIR,
    BLUE,
    CORAL,
    CREAM,
    GRAY,
    HANDBOOK_ASSETS,
    INK,
    MINT,
    NAVY,
    PALE,
    PEACH,
    SKY,
    SLIDE_H,
    SLIDE_W,
    TEAL,
    WHITE,
    add_picture_cover,
    add_round_rect,
    rgb,
)


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "Physical-Social-Games-Kids-Explainer-Arabic.pptx"
ARABIC_FONT = "Arial"


GAMES = [
    {
        "name": "مطاردة النبض",
        "tagline": "لعبة مطاردة ولمس آمنة",
        "players": "٤–٦",
        "movement": "نشطة",
        "image": "01-pulse-hunt.png",
        "goal": "يبقى العدّاؤون أحرارًا حتى ينتهي الوقت، بينما يحاول الصياد لمس جهاز كل عدّاء.",
        "action": "يضغط الصياد على جهاز ظهر العدّاء فقط. عندما يضيء الجهاز نعرف أن اللمسة حُسبت.",
        "win": "يفوز العدّاؤون إذا بقي واحد منهم حرًا، ويفوز الصياد إذا لمس الجميع.",
        "safety": "المس الجهاز فقط؛ ممنوع الدفع أو الإمساك أو العرقلة أو إسقاط أي لاعب.",
    },
    {
        "name": "سرقة الخزنة",
        "tagline": "اسرق الجواهر ثم احمِ خزنتك",
        "players": "٤–٦",
        "movement": "متوسطة",
        "image": "02-vault-heist.png",
        "goal": "اجمع الجواهر الرقمية وحاول منع الفريق الآخر من إفراغ خزنتك.",
        "action": "اضغط مطولًا على جهاز الخصم لسرقة جوهرة. بعدها يحميه درع لوقت قصير.",
        "win": "احمِ خزنتك وأنهِ الجولة بجواهر أكثر من الفريق الآخر.",
        "safety": "اضغط الجهاز بلطف، ولا تمسك اللاعب الآخر أو تسحبه أو تحاصره.",
    },
    {
        "name": "سلسلة الإشارة",
        "tagline": "تذكّر الترتيب وحافظ على السلسلة",
        "players": "٣–٦",
        "movement": "خفيفة",
        "image": "03-signal-chain.png",
        "goal": "تعاونوا لتذكّر ترتيب متزايد من اللاعبين والألوان.",
        "action": "راقبوا الأجهزة التي تضيء، ثم اضغط جهازك عندما يحين دورك.",
        "win": "أكملوا السلسلة كاملة قبل انتهاء الوقت.",
        "safety": "ابقَ في مكانك وانتظر دورك؛ السرعة هنا تأتي من التفكير.",
    },
    {
        "name": "رابط الحارس",
        "tagline": "شركاء سريّون يحمون بعضهم",
        "players": "٤–٦",
        "movement": "متوسطة",
        "image": "04-guardian-link.png",
        "goal": "أنجز المهام العلنية بينما تحافظ سرًا على سلامة شريكك.",
        "action": "إذا أُصيب شريكك، اضغط جهازه خلال ثلاث ثوانٍ لتصنع درع إنقاذ.",
        "win": "احمِ شريكك وأنجز أكبر عدد ممكن من مهام الفريق.",
        "safety": "الإنقاذ بلمس الجهاز فقط؛ ممنوع لمس الجسم أو سد الطريق.",
    },
    {
        "name": "نبض المناطق",
        "tagline": "سيطر على المناطق لفريقك",
        "players": "٤–٦",
        "movement": "متوسطة",
        "image": "05-territory-pulse.png",
        "goal": "اجعل مناطق أكثر على الأرض تضيء بلون فريقك.",
        "action": "اضغط مطولًا على جهاز المنطقة حتى يتغير لونه. المناطق المسيطر عليها تكسب نقاطًا.",
        "win": "يفوز الفريق الذي يسيطر على المناطق لأطول وقت.",
        "safety": "امشِ بين المناطق وضع كل جهاز في مكان لا يدوس عليه أحد.",
    },
    {
        "name": "مطابقة الصدى",
        "tagline": "شاهد النمط ثم انسخه بسرعة",
        "players": "٢–٦",
        "movement": "خفيفة",
        "image": "06-echo-match.png",
        "goal": "تذكّر نمط الألوان أو الأصوات الذي تعرضه الوحدة المركزية.",
        "action": "راقب جيدًا، ثم اضغط جهازك بالترتيب نفسه.",
        "win": "الأنماط الصحيحة والسريعة تكسب نقاطًا. يفوز اللاعب الأدق.",
        "safety": "ثبّت الأجهزة جيدًا واستخدم الضوء أو الاهتزاز إذا كان الصوت مزعجًا.",
    },
    {
        "name": "المهمة الصامتة",
        "tagline": "أنهِ مهمتك السرية دون كشفها",
        "players": "٣–٦",
        "movement": "خفيفة",
        "image": "07-silent-mission.png",
        "goal": "أنهِ مهمتك الخفية أثناء حديث أو تجمع عادي.",
        "action": "استخدم ضغطة قصيرة أو طويلة على جهازك لتسجيل كل خطوة بهدوء.",
        "win": "في النهاية تكشف الوحدة المركزية من أكمل مهمته السرية.",
        "safety": "يجب أن تكون المهمات لطيفة وآمنة وتحترم مساحة الجميع.",
    },
    {
        "name": "سباق التتابع",
        "tagline": "مرّر الضوء عبر الفريق كله",
        "players": "٣–٦",
        "movement": "خفيفة–متوسطة",
        "image": "08-relay-rush.png",
        "goal": "حرّك الإشارة المضيئة عبر الفريق بأسرع وقت ممكن.",
        "action": "عندما يضيء جهازك اضغطه لإرسال الإشارة إلى اللاعب التالي.",
        "win": "أكمل المسار بأقل وقت؛ الأخطاء تضيف ثوانٍ إضافية.",
        "safety": "يبقى الجهاز مثبتًا أو على الطاولة؛ مرّر الإشارة لا الجهاز نفسه.",
    },
    {
        "name": "دائرة التاج",
        "tagline": "احتفظ بالتاج واربح تحديات السرعة",
        "players": "٤–٦",
        "movement": "متوسطة",
        "image": "09-crown-circuit.png",
        "goal": "احتفظ بالتاج الرقمي لتجمع أكبر عدد من نقاط التاج.",
        "action": "تحدَّ حامل التاج، ابتعد عنه، ثم اضغط جهازك بعد ظهور الوميض فقط.",
        "win": "احتفظ بالتاج واربح التحديات واصل إلى النقاط المطلوبة أولًا.",
        "safety": "امشِ في الأماكن الصغيرة، والمس الجهاز فقط واحترم درع الفائز.",
    },
    {
        "name": "شفرة الساعي",
        "tagline": "أخفِ الحزمة الرقمية ومرّرها سرًا",
        "players": "٥–٦",
        "movement": "خفيفة",
        "image": "10-courier-code.png",
        "goal": "يمرّر فريق الساعي الحزمة سرًا ثلاث مرات ثم يخرج بها.",
        "action": "يمسك حامل الحزمة جهاز زميله، ثم يضغط الزميل جهازه ليستلمها.",
        "win": "يفوز السعاة بالخروج، ويفوز المعترض إذا فحص حامل الحزمة الحقيقي.",
        "safety": "ممنوع تفتيش الملابس أو إمساك الأيدي أو إجبار أحد على كشف دليل.",
    },
    {
        "name": "القفل المزدوج",
        "tagline": "حلوا الألغاز واضغطوا معًا",
        "players": "٣–٦",
        "movement": "خفيفة",
        "image": "11-double-lock.png",
        "goal": "افتحوا ستة أقفال رقمية بحل كل لغز كفريق.",
        "action": "اختاروا الجهازين أو الأجهزة الثلاثة الصحيحة واضغطوها في الوقت نفسه.",
        "win": "افتحوا الأقفال الستة قبل أن يصل المؤقت إلى الصفر.",
        "safety": "ثبّتوا الأجهزة واتركوا لكل لاعب مساحة واضحة للوصول إليها.",
    },
    {
        "name": "سباق السوق",
        "tagline": "تبادل الموارد وأكمل العقود",
        "players": "٤–٦",
        "movement": "جلوس",
        "image": "12-market-rush.png",
        "goal": "تبادل الشرارة والقماش والوقت لإكمال عقود عامة مفيدة.",
        "action": "كوّن العرض بجهازك. لا تتم الصفقة إلا عندما يؤكد اللاعبان معًا.",
        "win": "العقود المكتملة تكسب سمعة، ويفوز صاحب أعلى سمعة.",
        "safety": "كل شيء خيالي؛ لا مال حقيقي ولا مراهنات ولا وعود أو جوائز مالية.",
    },
    {
        "name": "شبكة الطاقة",
        "tagline": "حافظ على طاقة المدينة وبرودة المولدات",
        "players": "٣–٦",
        "movement": "خفيفة",
        "image": "13-power-grid.png",
        "goal": "طابق حاجة المدينة المتغيرة للطاقة دون أن تسخن المولدات كثيرًا.",
        "action": "اضغط لزيادة الطاقة، واضغط مطولًا للتبريد، واربط جهازين للإصلاح.",
        "win": "حافظ على مؤشر الاستقرار فوق الصفر حتى تنتهي فترة العمل.",
        "safety": "هذه لعبة طاولة؛ اجعل الأجهزة مسطحة وجافة وسهلة الوصول.",
    },
    {
        "name": "منافسو الإيقاع",
        "tagline": "اضغط مع النغمة وابنِ إيقاع الفريق",
        "players": "٢–٦",
        "movement": "خفيفة",
        "image": "14-rhythm-rivals.png",
        "goal": "اضغط نغماتك في اللحظة الصحيحة وحافظ على سلسلة إيقاع ناجحة.",
        "action": "يضيء جهازك قبل النغمة. اضغط عند الإشارة، ونغمات الفريق تحتاج الجميع معًا.",
        "win": "تفوز أعلى نقاط إيقاع، أو يحقق الفريق كله هدف الدقة.",
        "safety": "استخدم صوتًا مريحًا، ولا تضرب الجهاز بالأثاث أو الأشخاص.",
    },
    {
        "name": "تبديل المسار",
        "tagline": "وجّه الإشارة بينما تتغير القواعد",
        "players": "٤–٦",
        "movement": "جلوس",
        "image": "15-switchback.png",
        "goal": "حافظ على حركة الإشارة نحو لاعب صحيح دون ارتكاب خطأ.",
        "action": "الضغط القصير والطويل والمزدوج يرسل الإشارة بطرق مختلفة، ثم تتغير القاعدة.",
        "win": "اصمدوا حتى نهاية الجولة، أو احتفظ بأكبر عدد من الفرص في الوضع التنافسي.",
        "safety": "ابقَ جالسًا واضغط جهازك؛ لا ترمِ الجهاز ولا تمرره بيدك.",
    },
    {
        "name": "تردد الشبح",
        "tagline": "اختبر الأزواج واجمع الأدلة واكتشف المتخفي",
        "players": "٤–٦",
        "movement": "خفيفة",
        "image": "16-ghost-frequency.png",
        "goal": "استخدم عددًا قليلًا من اختبارات الأزواج لمعرفة اللاعب المتخفي.",
        "action": "اختبر جهازين، وقارن نتيجة متشابه أو مختلف، ثم صوّت سرًا.",
        "win": "يفوز المحققون باكتشاف المتخفي، ويفوز المتخفي إذا بقي مجهولًا.",
        "safety": "الاتهامات داخل اللعبة فقط؛ تكلم بلطف وأشرك الجميع بعد انتهائها.",
    },
]


def arabic_number(number):
    return str(number).translate(str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩"))


def set_arabic_run(run, font, size, color, bold):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "ar-AE")
    rpr.set("altLang", "en-US")
    for tag in ("a:cs", "a:ea"):
        node = OxmlElement(tag)
        node.set("typeface", font)
        rpr.append(node)


def add_ar_text(slide, text, x, y, w, h, size=20, color=INK, bold=False,
                align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.TOP, margin=0.06,
                line_spacing=1.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    frame._txBody.bodyPr.set("rtlCol", "1")
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.line_spacing = line_spacing
    paragraph._p.get_or_add_pPr().set("rtl", "1")
    run = paragraph.add_run()
    run.text = text
    set_arabic_run(run, ARABIC_FONT, size, color, bold)
    return box


def add_top_bar(slide, section, number=None):
    add_round_rect(slide, 0, 0, SLIDE_W, 0.62, NAVY, MSO_SHAPE.RECTANGLE)
    add_ar_text(slide, "الألعاب الاجتماعية الحركية", 8.35, 0.14, 4.55, 0.28,
                11, WHITE, True, margin=0)
    label = section if number is None else f"{section}  /  {arabic_number(f'{number:02d}')}"
    add_ar_text(slide, label, 0.42, 0.14, 4.7, 0.28, 10.5, "7FDBD4", True,
                align=PP_ALIGN.LEFT, margin=0)


def add_pill(slide, text, x, y, w, fill=MINT, color=TEAL):
    add_round_rect(slide, x, y, w, 0.42, fill, line=color, line_width=0.8)
    add_ar_text(slide, text, x + 0.06, y + 0.075, w - 0.12, 0.24, 9.2, color, True,
                align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0)


def add_info_card(slide, label, body, x, y, w, fill, accent):
    add_round_rect(slide, x, y, w, 1.0, fill)
    add_round_rect(slide, x + w - 0.10, y, 0.10, 1.0, accent, MSO_SHAPE.RECTANGLE)
    add_ar_text(slide, label, x + 0.18, y + 0.10, w - 0.39, 0.21, 9.4, accent, True, margin=0)
    add_ar_text(slide, body, x + 0.18, y + 0.33, w - 0.39, 0.59, 12.7, INK,
                margin=0, line_spacing=0.94)


def add_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(NAVY)
    add_round_rect(slide, 7.43, 0.55, 5.35, 6.4, "163954")
    add_ar_text(slide, "١٦ لعبة.\nصندوق سحري واحد.", 7.88, 1.0, 4.45, 1.45, 29, WHITE, True,
                line_spacing=0.9)
    add_ar_text(slide, "دليل بصري يجعل كل لعبة سهلة الفهم.", 7.88, 2.72, 4.45, 0.92,
                19, "CFE8E8")
    add_round_rect(slide, 8.63, 4.05, 3.7, 0.58, AMBER)
    add_ar_text(slide, "شاهد  •  اضغط  •  العب", 8.76, 4.17, 3.45, 0.28, 14,
                NAVY, True, align=PP_ALIGN.CENTER, margin=0)
    add_ar_text(slide, "نسخة شرح مبسطة للأطفال\nيبقى توجيه النموذج الحالي: تحت الإشراف لعمر ١٤+.",
                7.88, 5.25, 4.45, 0.82, 12, "AFC3D4")
    add_round_rect(slide, 0.55, 0.55, 6.58, 6.4, WHITE)
    add_picture_cover(slide, HANDBOOK_ASSETS / "kit-concept.png", 0.70, 0.72, 6.28, 5.45,
                      "صورة تصورية لستة أجهزة لعب ووحدة مركزية وحقيبة حمل")
    add_ar_text(slide, "ستة أجهزة + وحدة مركزية = ألعاب كثيرة", 0.96, 6.26, 5.7, 0.38,
                16, NAVY, True, align=PP_ALIGN.CENTER, margin=0)


def add_how_it_works_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(PALE)
    add_top_bar(slide, "كيف تعمل؟")
    add_ar_text(slide, "الأزرار نفسها تتحول إلى ألعاب مختلفة", 4.0, 0.88, 8.75, 0.55, 27, NAVY, True)
    add_ar_text(slide, "الجهاز يخبر الوحدة المركزية بما فعلته، وقواعد اللعبة تحدد معنى الحركة.",
                0.72, 1.44, 12.0, 0.45, 16, GRAY)
    cards = [
        ("١", "ارتدِ الجهاز أو ضعه", "ثبّته بأمان، أمسكه بيدك، أو ضعه ثابتًا على الطاولة.", TEAL, MINT),
        ("٢", "راقب الإشارة", "توضح الوحدة والجهاز دورك وما يحدث في اللعبة.", CORAL, PEACH),
        ("٣", "استخدم حركة واحدة", "ضغطة قصيرة أو طويلة أو مزدوجة، أو اضغطوا معًا.", AMBER, CREAM),
    ]
    for i, (num, title, body, accent, fill) in enumerate(cards):
        x = 8.96 - i * 4.18
        add_round_rect(slide, x, 2.18, 3.75, 3.75, WHITE, line="D4DEE7", line_width=1)
        add_round_rect(slide, x + 1.36, 2.48, 1.02, 1.02, accent, MSO_SHAPE.OVAL)
        add_ar_text(slide, num, x + 1.36, 2.66, 1.02, 0.45, 24, WHITE, True,
                    align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0)
        add_ar_text(slide, title, x + 0.34, 3.82, 3.07, 0.42, 15, accent, True,
                    align=PP_ALIGN.CENTER, margin=0)
        add_ar_text(slide, body, x + 0.39, 4.40, 2.97, 0.90, 15.5, INK,
                    align=PP_ALIGN.CENTER, margin=0, line_spacing=1.02)
    add_round_rect(slide, 0.62, 6.36, 12.05, 0.6, NAVY)
    add_ar_text(slide, "تفحص الوحدة المركزية كل ضغطة، وتحفظ النقاط، وتخبر كل جهاز بما يعرضه.",
                0.95, 6.49, 11.4, 0.30, 14, WHITE, True,
                align=PP_ALIGN.CENTER, margin=0)


def add_safety_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(PALE)
    add_top_bar(slide, "السلامة أولًا")
    add_ar_text(slide, "سهولة الشرح لا تعني أن اللعبة جاهزة لعمر ١٠ سنوات",
                2.5, 0.86, 10.25, 0.58, 27, NAVY, True)
    add_picture_cover(slide, HANDBOOK_ASSETS / "gameplay-villa.png", 5.55, 1.62, 7.2, 4.63,
                      "بالغون يلعبون بأمان بالأجهزة في حديقة خالية من العوائق")
    rules = [
        ("الخطة الحالية", "النموذج الأولي تحت الإشراف ومخصص لعمر ١٤+ حتى يكتمل اختبار السلامة.", CORAL, PEACH),
        ("المس الجهاز فقط", "ممنوع الدفع والإمساك والعرقلة ولمس الوجه أو الرقبة.", TEAL, MINT),
        ("جهّز المكان", "ابتعد عن الطرق والسلالم والمسابح والمطابخ والسيارات والأشياء القابلة للكسر.", AMBER, CREAM),
        ("توقف فورًا", "توقف عند الألم أو الحرارة أو تلف الجهاز أو ارتخاء الحزام أو السلوك غير الآمن.", BLUE, SKY),
    ]
    for i, (title, body, accent, fill) in enumerate(rules):
        y = 1.62 + i * 1.18
        add_round_rect(slide, 0.63, y, 4.65, 0.98, fill)
        add_ar_text(slide, title, 0.84, y + 0.10, 4.13, 0.23, 10, accent, True, margin=0)
        add_ar_text(slide, body, 0.84, y + 0.36, 4.13, 0.49, 12.7, INK,
                    margin=0, line_spacing=0.94)
    add_ar_text(slide, "يفحص المشرف المكان ويشرح القواعد قبل كل جولة.", 0.75, 6.48, 11.9, 0.34,
                15, NAVY, True, align=PP_ALIGN.CENTER, margin=0)


def add_game_slide(prs, game, index):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(PALE)
    add_top_bar(slide, "مكتبة الألعاب", index)
    add_ar_text(slide, game["name"], 6.68, 0.82, 6.10, 0.55, 28, NAVY, True)
    add_ar_text(slide, game["tagline"], 6.38, 1.34, 6.40, 0.34, 14, GRAY)
    add_pill(slide, f"لاعبون: {game['players']}", 0.55, 0.88, 1.55)
    add_pill(slide, game["movement"], 2.24, 0.88, 1.42, CREAM, AMBER)
    add_pill(slide, "مركز واحد", 3.80, 0.88, 1.35, SKY, BLUE)

    add_round_rect(slide, 5.28, 1.82, 7.55, 4.78, WHITE, line="D4DEE7", line_width=1)
    add_picture_cover(slide, ASSET_DIR / game["image"], 5.39, 1.93, 7.33, 4.56,
                      f"شرح بصري للعبة {game['name']}")

    add_info_card(slide, "هدف اللعبة", game["goal"], 0.55, 1.82, 4.48, MINT, TEAL)
    add_info_card(slide, "ماذا تفعل؟", game["action"], 0.55, 3.02, 4.48, PEACH, CORAL)
    add_info_card(slide, "كيف تفوز؟", game["win"], 0.55, 4.22, 4.48, CREAM, AMBER)

    add_round_rect(slide, 0.55, 5.46, 4.48, 1.14, NAVY)
    add_ar_text(slide, "لعب آمن", 3.66, 5.59, 1.08, 0.21, 9.5, "7FDBD4", True, margin=0)
    add_ar_text(slide, game["safety"], 0.80, 5.87, 3.93, 0.55, 12.5, WHITE,
                margin=0, line_spacing=0.93)
    add_ar_text(slide, arabic_number(f"{index:02d}"), 12.20, 6.82, 0.48, 0.23,
                9, GRAY, True, margin=0)
    add_ar_text(slide, "دليل تعلّم بصري • القواعد مبسطة للشرح", 0.62, 6.82, 4.55, 0.23,
                9, GRAY, align=PP_ALIGN.LEFT, margin=0)


def add_chooser_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = rgb(PALE)
    add_top_bar(slide, "اختر لعبتك")
    add_ar_text(slide, "ما نوع اللعبة التي تريدها اليوم؟", 3.0, 0.88, 9.75, 0.56, 27, NAVY, True)
    groups = [
        ("حركة ومطاردة", "مطاردة النبض\nسرقة الخزنة\nرابط الحارس\nنبض المناطق\nدائرة التاج", CORAL, PEACH),
        ("فكّروا معًا", "سلسلة الإشارة\nسباق التتابع\nالقفل المزدوج\nشبكة الطاقة", TEAL, MINT),
        ("تحدث وخادع", "المهمة الصامتة\nشفرة الساعي\nسباق السوق\nتردد الشبح", AMBER, CREAM),
        ("أيدٍ سريعة", "مطابقة الصدى\nمنافسو الإيقاع\nتبديل المسار", BLUE, SKY),
    ]
    for i, (title, games, accent, fill) in enumerate(groups):
        x = 9.90 - i * 3.15
        add_round_rect(slide, x, 1.83, 2.83, 4.6, WHITE, line="D4DEE7", line_width=1)
        add_round_rect(slide, x, 1.83, 2.83, 0.72, accent, MSO_SHAPE.RECTANGLE)
        add_ar_text(slide, title, x + 0.15, 2.04, 2.53, 0.29, 13, WHITE, True,
                    align=PP_ALIGN.CENTER, margin=0)
        add_ar_text(slide, games, x + 0.30, 2.87, 2.23, 2.85, 17, NAVY, True,
                    align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0,
                    line_spacing=1.18)
        add_round_rect(slide, x + 0.96, 5.83, 0.90, 0.22, fill)
    add_round_rect(slide, 1.47, 6.70, 10.38, 0.45, NAVY)
    add_ar_text(slide, "صندوق واحد يمكن أن يصبح مطاردة أو لغزًا أو حفلة أو لعبة إيقاع أو استراتيجية.",
                1.70, 6.80, 9.9, 0.24, 13, WHITE, True,
                align=PP_ALIGN.CENTER, margin=0)


def build():
    for game in GAMES:
        path = ASSET_DIR / game["image"]
        if not path.exists():
            raise FileNotFoundError(path)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    props = prs.core_properties
    props.title = "الألعاب الاجتماعية الحركية - دليل بصري مبسط"
    props.subject = "شرح عربي مبسط لجميع أفكار الألعاب الست عشرة"
    props.author = "مشروع الألعاب الاجتماعية الحركية"
    props.keywords = "ألعاب حركية، دليل بصري، ألعاب اجتماعية، أجهزة قابلة للارتداء"
    props.comments = "الشرح مبسط للفهم. يبقى توجيه النموذج الأولي الحالي: تحت الإشراف لعمر ١٤+."

    add_title_slide(prs)
    add_how_it_works_slide(prs)
    add_safety_slide(prs)
    for index, game in enumerate(GAMES, 1):
        add_game_slide(prs, game, index)
    add_chooser_slide(prs)

    prs.save(OUTPUT)
    print(f"Saved {OUTPUT} with {len(prs.slides)} slides")


if __name__ == "__main__":
    build()
