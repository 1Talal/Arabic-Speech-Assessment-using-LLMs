const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer, PageBreak
} = require('docx');

// =========================================================
// Helpers - Arabic RTL text
// =========================================================
const AR = (text, opts = {}) => new TextRun({
  text,
  rightToLeft: true,
  font: 'Arial',
  ...opts,
});

const P = (text, opts = {}) => new Paragraph({
  bidirectional: true,
  alignment: opts.alignment || AlignmentType.RIGHT,
  spacing: { after: 120, ...opts.spacing },
  ...opts.paragraph,
  children: Array.isArray(text)
    ? text
    : [AR(text, opts.run || {})],
});

const H1 = (text) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 200 },
  children: [AR(text, { bold: true, size: 32, color: '1F3864' })],
});

const H2 = (text) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 160 },
  children: [AR(text, { bold: true, size: 26, color: '2E75B6' })],
});

const bulletItem = (text) => new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.RIGHT,
  numbering: { reference: 'bullets', level: 0 },
  spacing: { after: 80 },
  children: [AR(text)],
});

// =========================================================
// Table helpers
// =========================================================
const BORDER = { style: BorderStyle.SINGLE, size: 6, color: '8FAADC' };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

const cell = (text, opts = {}) => new TableCell({
  borders: BORDERS,
  width: { size: opts.width, type: WidthType.DXA },
  shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 120, right: 120 },
  children: [new Paragraph({
    bidirectional: true,
    alignment: opts.align || AlignmentType.CENTER,
    children: [AR(String(text), { bold: opts.bold, color: opts.color, size: 22 })],
  })],
});

// =========================================================
// Data from results CSV
// =========================================================
const transcript = 'السلام عليكم أنا طرال أحب أن أتعلم اللغة الأرضية';

const results = [
  {
    model: 'GPT-4o Mini', provider: 'OpenAI',
    pronunciation: 6, fluency: 5, grammar: 6, vocabulary: 5, overall: 5, cefr: 'A1',
    strengths: ['استخدام التحية بشكل صحيح', 'التعبير عن الفكرة الأساسية'],
    weaknesses: ['عدم اكتمال الجملة', 'استخدام مفردات بسيطة جداً'],
    recommendation: 'العمل على بناء جمل كاملة وتوسيع المفردات.',
    feedback: 'الطالب استخدم تحية صحيحة، لكن الجملة غير مكتملة مما يؤثر على الفهم. يجب على الطالب التركيز على بناء جمل كاملة واستخدام مفردات متنوعة لتحسين مستوى اللغة.',
  },
  {
    model: 'Qwen 2.5 72B', provider: 'Alibaba',
    pronunciation: 5, fluency: 4, grammar: 3, vocabulary: 4, overall: 4, cefr: 'A1',
    strengths: ['استخدام التحية بشكل صحيح'],
    weaknesses: ['خلل في الجملة', 'ضعف في بناء الجملة'],
    recommendation: 'التدرب على بناء الجمل البسيطة بشكل صحيح',
    feedback: 'الطالب استخدم التحية بشكل صحيح، لكن هناك خلل في بناء الجملة. يُنصح بالتدرب على بناء الجمل البسيطة بشكل صحيح لتحسين مستوى اللغة.',
  },
  {
    model: 'Llama 3.3 70B', provider: 'Meta',
    pronunciation: 4, fluency: 3, grammar: 2, vocabulary: 3, overall: 3, cefr: 'A1',
    strengths: ['القدرة على بدء الحديث'],
    weaknesses: ['نقص في القواعد النحوية', 'صعوبة في التعبير عن الأفكار'],
    recommendation: 'يُوصى بالتركيز على تعلم القواعد النحوية الأساسية وتحسين مهارات التعبير',
    feedback: 'يظهر الطالب صعوبة في بناء الجمل بشكل صحيح ويتعذر عليه التعبير عن أفكاره بفعالية. يحتاج إلى مزيد من الممارسة في استخدام القواعد النحوية والاهتمام بتحسين مستوى المفردات.',
  },
  {
    model: 'Mistral Large', provider: 'Mistral',
    pronunciation: 6, fluency: 4, grammar: 3, vocabulary: 4, overall: 4, cefr: 'A1',
    strengths: ['النطق واضح نسبياً في الكلمات البسيطة', 'محاولة استخدام تحية شائعة'],
    weaknesses: ['خطأ نحوي في تركيب الجملة', 'قلة المفردات', 'عدم طلاقة في الكلام'],
    recommendation: 'التركيز على تعلم التراكيب الأساسية للجمل والتمرن على المحادثات القصيرة',
    feedback: 'يتميز الطالب بقدرته على نطق الكلمات البسيطة بوضوح، لكنه يواجه صعوبات في بناء الجمل بشكل صحيح. المفردات المستخدمة محدودة، ولا يظهر طلاقة في الكلام. المستوى الحالي يقارب A1، ويحتاج إلى ممارسة أكثر.',
  },
];

const best = results.reduce((a, b) => (a.overall >= b.overall ? a : b));
const avgOverall = (results.reduce((s, r) => s + r.overall, 0) / results.length).toFixed(1);
const avgPron = (results.reduce((s, r) => s + r.pronunciation, 0) / results.length).toFixed(1);
const avgFlu  = (results.reduce((s, r) => s + r.fluency, 0) / results.length).toFixed(1);
const avgGra  = (results.reduce((s, r) => s + r.grammar, 0) / results.length).toFixed(1);
const avgVoc  = (results.reduce((s, r) => s + r.vocabulary, 0) / results.length).toFixed(1);

// =========================================================
// Build the document content
// =========================================================
const children = [];

// === Cover ===
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  spacing: { before: 2400, after: 200 },
  children: [AR('تقرير مشروع', { bold: true, size: 36, color: '2E75B6' })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  spacing: { after: 400 },
  children: [AR('تقييم النطق العربي بالذكاء الاصطناعي', { bold: true, size: 56, color: '1F3864' })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [AR('Arabic Speech Assessment with Multi-LLM Comparison', { italics: true, size: 26, color: '595959' })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  spacing: { before: 2400, after: 200 },
  children: [AR('إعداد: مجاهد', { size: 28, bold: true })],
}));
children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  children: [AR('مايو 2026', { size: 24, color: '595959' })],
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// === Executive Summary ===
children.push(H1('الملخص التنفيذي'));
children.push(P(
  'يهدف هذا المشروع إلى بناء نظام آلي لتقييم مستوى اللغة العربية المنطوقة باستخدام نماذج الذكاء الاصطناعي الكبيرة (LLMs). ' +
  'يقوم النظام بتحويل التسجيل الصوتي إلى نص باستخدام نموذج Whisper، ثم يرسل النص إلى عدة نماذج ذكاء اصطناعي عبر منصة OpenRouter ' +
  'للحصول على تقييمات متعددة تشمل النطق والطلاقة والقواعد النحوية والمفردات، وتحديد المستوى وفق الإطار المرجعي الأوروبي المشترك (CEFR).'
));
children.push(P(
  'تم اختبار المشروع على عينة نطق قصيرة وقورنت نتائج أربعة نماذج: GPT-4o Mini و Qwen 2.5 72B و Llama 3.3 70B و Mistral Large. ' +
  `أعطى نموذج ${best.model} أعلى تقييم بدرجة عامة ${best.overall} من 10، وكان متوسط الدرجة العامة بين جميع النماذج ${avgOverall} من 10. ` +
  'اتفقت جميع النماذج على أن مستوى المتحدث في هذه العينة هو A1 وفق تصنيف CEFR.'
));

// === Introduction ===
children.push(H1('مقدمة المشروع'));
children.push(P(
  'تواجه عملية تقييم اللغة العربية المنطوقة تحديات تقليدية متعددة، أبرزها الاعتماد على المعلم البشري الذي قد يستهلك وقتاً طويلاً ' +
  'وقد يختلف تقييمه من شخص لآخر. كما أن قلة الأدوات المتخصصة في تقييم العربية مقارنة بالإنجليزية تجعل عملية تطوير منصات تعليمية ذكية ' +
  'باللغة العربية أكثر صعوبة.'
));
children.push(P(
  'يقدم هذا المشروع حلاً يعتمد على نماذج الذكاء الاصطناعي الحديثة لتقديم تقييم آلي شامل لمستوى الطالب في اللغة العربية. ' +
  'يتميز النظام بإمكانية مقارنة استجابات أكثر من نموذج في نفس الوقت، مما يساعد على تحديد النموذج الأنسب لأغراض التعليم العربي.'
));

children.push(H2('أهداف المشروع'));
children.push(bulletItem('بناء واجهة سهلة الاستخدام لتقييم النطق العربي بالذكاء الاصطناعي.'));
children.push(bulletItem('تحويل الصوت العربي إلى نص باستخدام نموذج Whisper.'));
children.push(bulletItem('تقييم النص العربي باستخدام عدة نماذج LLM متوازية عبر OpenRouter.'));
children.push(bulletItem('مقارنة أداء النماذج المختلفة من حيث الدقة والاتساق.'));
children.push(bulletItem('تقديم نتائج مرئية واضحة تشمل الرسوم البيانية والتوصيات.'));

// === Methodology ===
children.push(H1('المنهجية والبنية التقنية'));

children.push(H2('الأدوات والتقنيات المستخدمة'));
children.push(bulletItem('Streamlit: لبناء واجهة المستخدم التفاعلية مع دعم اللغة العربية واتجاه RTL.'));
children.push(bulletItem('OpenAI Whisper / Groq Whisper: لتحويل التسجيل الصوتي العربي إلى نص.'));
children.push(bulletItem('OpenRouter API: لاستدعاء عدة نماذج LLM من مزودين مختلفين عبر واجهة موحدة.'));
children.push(bulletItem('Plotly: لإنشاء الرسوم البيانية التفاعلية (Bar Chart و Radar Chart).'));
children.push(bulletItem('Pandas: لمعالجة وتصدير النتائج بصيغة CSV.'));
children.push(bulletItem('FFmpeg: لتحويل ملفات الصوت إلى الصيغة المطلوبة.'));

children.push(H2('مراحل العمل'));
children.push(P([AR('المرحلة الأولى - الإدخال: ', { bold: true }), AR('يقوم المستخدم برفع ملف صوتي بالعربية، أو يدخل النص مباشرة.')]));
children.push(P([AR('المرحلة الثانية - التفريغ الصوتي: ', { bold: true }), AR('يتم تحويل الصوت إلى نص عربي باستخدام Whisper.')]));
children.push(P([AR('المرحلة الثالثة - التقييم المتوازي: ', { bold: true }), AR('يُرسل النص إلى عدة نماذج LLM عبر OpenRouter بشكل متوازٍ لتسريع المعالجة.')]));
children.push(P([AR('المرحلة الرابعة - تحليل وعرض النتائج: ', { bold: true }), AR('تُجمع نتائج النماذج وتُعرض في شكل جداول ورسوم بيانية وتفاصيل لكل نموذج.')]));

children.push(H2('معايير التقييم'));
children.push(P('يقيِّم كل نموذج النص وفق خمسة معايير رئيسية، كل منها بدرجة من صفر إلى عشرة:'));
children.push(bulletItem('النطق (Pronunciation): مدى وضوح الكلمات وصحة لفظها.'));
children.push(bulletItem('الطلاقة (Fluency): مدى تدفق الكلام بسلاسة دون توقف.'));
children.push(bulletItem('القواعد النحوية (Grammar): صحة التركيب النحوي للجمل.'));
children.push(bulletItem('المفردات (Vocabulary): تنوع وثراء المفردات المستخدمة.'));
children.push(bulletItem('الدرجة العامة (Overall): تقييم شامل يأخذ بعين الاعتبار جميع المعايير.'));
children.push(P('بالإضافة إلى ذلك، يحدد كل نموذج المستوى وفق تصنيف CEFR من A1 (مبتدئ) إلى C2 (متمكن).'));

children.push(new Paragraph({ children: [new PageBreak()] }));

// === Sample ===
children.push(H1('عينة الاختبار'));
children.push(P('تم اختبار النظام على تسجيل صوتي قصير بالعربية، وكان النص الناتج من عملية التفريغ كالتالي:'));
children.push(new Paragraph({
  bidirectional: true,
  alignment: AlignmentType.CENTER,
  spacing: { before: 200, after: 200 },
  shading: { fill: 'EDEDED', type: ShadingType.CLEAR },
  border: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
  children: [AR(`"${transcript}"`, { italics: true, size: 28, color: '1F3864', bold: true })],
}));
children.push(P(
  'يُلاحظ من النص وجود خطأ واضح: الكلمة "الأرضية" بدلاً من "العربية"، وهذا قد يكون نتيجة خطأ في التفريغ الصوتي أو خطأ في النطق من المتحدث. ' +
  'هذه نقطة مهمة سوف نناقشها لاحقاً عند تحليل أداء النماذج.'
));

// === Results Table ===
children.push(H1('نتائج التقييم - مقارنة شاملة'));
children.push(P('يلخص الجدول التالي تقييم كل نموذج على معايير التقييم الخمسة بالإضافة إلى المستوى المحدد:'));

const totalWidth = 9360;
const colWidths = [1760, 1100, 1100, 1100, 1100, 1100, 1200, 900];
// Sum: 1760+1100*5+1200+900 = 1760+5500+1200+900 = 9360 ✓

children.push(new Table({
  width: { size: totalWidth, type: WidthType.DXA },
  columnWidths: colWidths,
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        cell('النموذج',     { width: colWidths[0], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('المزود',       { width: colWidths[1], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('النطق',        { width: colWidths[2], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('الطلاقة',      { width: colWidths[3], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('النحو',        { width: colWidths[4], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('المفردات',     { width: colWidths[5], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('الدرجة العامة', { width: colWidths[6], shading: '1F3864', color: 'FFFFFF', bold: true }),
        cell('CEFR',         { width: colWidths[7], shading: '1F3864', color: 'FFFFFF', bold: true }),
      ],
    }),
    ...results.map((r, i) => new TableRow({
      children: [
        cell(r.model,         { width: colWidths[0], shading: i % 2 ? 'FFFFFF' : 'F2F2F2', bold: true }),
        cell(r.provider,      { width: colWidths[1], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
        cell(r.pronunciation, { width: colWidths[2], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
        cell(r.fluency,       { width: colWidths[3], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
        cell(r.grammar,       { width: colWidths[4], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
        cell(r.vocabulary,    { width: colWidths[5], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
        cell(r.overall,       { width: colWidths[6], shading: r === best ? 'FFD966' : (i % 2 ? 'FFFFFF' : 'F2F2F2'), bold: r === best }),
        cell(r.cefr,          { width: colWidths[7], shading: i % 2 ? 'FFFFFF' : 'F2F2F2' }),
      ],
    })),
    new TableRow({
      children: [
        cell('المتوسط', { width: colWidths[0], shading: '8FAADC', bold: true, color: 'FFFFFF' }),
        cell('-',       { width: colWidths[1], shading: '8FAADC', color: 'FFFFFF' }),
        cell(avgPron,   { width: colWidths[2], shading: '8FAADC', color: 'FFFFFF', bold: true }),
        cell(avgFlu,    { width: colWidths[3], shading: '8FAADC', color: 'FFFFFF', bold: true }),
        cell(avgGra,    { width: colWidths[4], shading: '8FAADC', color: 'FFFFFF', bold: true }),
        cell(avgVoc,    { width: colWidths[5], shading: '8FAADC', color: 'FFFFFF', bold: true }),
        cell(avgOverall,{ width: colWidths[6], shading: '8FAADC', color: 'FFFFFF', bold: true }),
        cell('A1',      { width: colWidths[7], shading: '8FAADC', color: 'FFFFFF', bold: true }),
      ],
    }),
  ],
}));

children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun('')] }));
children.push(P(
  `النموذج الأفضل أداءً هو ${best.model} بدرجة عامة ${best.overall} من 10. ` +
  `بينما النموذج الأقل تقييماً هو Llama 3.3 70B بدرجة 3 فقط، مما يُظهر تباينًا واضحاً بين النماذج رغم اتفاقها على المستوى A1.`
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// === Per-model analysis ===
children.push(H1('تحليل تفصيلي لكل نموذج'));

results.forEach((r) => {
  children.push(H2(`${r.model} (${r.provider})`));

  // Score line
  children.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { after: 120 },
    children: [
      AR('الدرجة العامة: ', { bold: true }),
      AR(`${r.overall}/10`, { bold: true, color: r === best ? 'C00000' : '1F3864' }),
      AR(' | مستوى CEFR: ', { bold: true }),
      AR(r.cefr, { bold: true, color: '1F3864' }),
    ],
  }));

  // Strengths
  children.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { before: 120, after: 80 },
    children: [AR('نقاط القوة:', { bold: true, color: '548235', size: 24 })],
  }));
  r.strengths.forEach(s => children.push(bulletItem(s)));

  // Weaknesses
  children.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { before: 120, after: 80 },
    children: [AR('نقاط الضعف:', { bold: true, color: 'C00000', size: 24 })],
  }));
  r.weaknesses.forEach(w => children.push(bulletItem(w)));

  // Recommendation
  children.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { before: 120, after: 80 },
    children: [AR('التوصية:', { bold: true, color: 'BF8F00', size: 24 })],
  }));
  children.push(P(r.recommendation));

  // Feedback
  children.push(new Paragraph({
    bidirectional: true, alignment: AlignmentType.RIGHT,
    spacing: { before: 120, after: 80 },
    children: [AR('التقييم التفصيلي:', { bold: true, color: '2E75B6', size: 24 })],
  }));
  children.push(P(r.feedback, { run: { italics: true } }));
});

children.push(new Paragraph({ children: [new PageBreak()] }));

// === Key findings ===
children.push(H1('النتائج الرئيسية والملاحظات'));

children.push(H2('1. تباين تقييمات النماذج'));
children.push(P(
  'رغم أن جميع النماذج اتفقت على أن المستوى هو A1، إلا أنها أعطت درجات مختلفة في المعايير الفرعية. ' +
  'الفرق بين أعلى تقييم (5/10 من GPT-4o Mini) وأقل تقييم (3/10 من Llama 3.3 70B) يصل إلى 40%، مما يشير إلى ' +
  'الحاجة إلى مرجعية موحدة عند بناء نظام تقييم آلي قابل للاعتماد.'
));

children.push(H2('2. أداء GPT-4o Mini'));
children.push(P(
  'كان GPT-4o Mini الأكثر اعتدالاً وتفصيلاً في التقييم. أعطى درجات أعلى بشكل ملحوظ في القواعد والمفردات، ' +
  'وقدم ملاحظات بناءة ومحددة دون مبالغة في انتقاد المتعلم المبتدئ.'
));

children.push(H2('3. حساسية النماذج للأخطاء'));
children.push(P(
  'الكلمة "الأرضية" بدلاً من "العربية" أربكت بعض النماذج. حيث رصدها Llama 3.3 70B كنقطة ضعف، ' +
  'بينما تجاهلها GPT-4o Mini ولم يشير إليها صراحةً. هذا يعكس تفاوت دقة النماذج في رصد الأخطاء الدلالية الدقيقة في العربية.'
));

children.push(H2('4. اتفاق على مستوى CEFR'));
children.push(P(
  'جميع النماذج الأربعة اتفقت على تصنيف المستوى كـ A1 (مبتدئ)، وهذا يعزز موثوقية تحديد المستوى الإجمالي ' +
  'حتى عندما تختلف الدرجات الرقمية. الاتفاق على المستوى مع اختلاف الدرجات قد يكون مؤشراً مهماً للأنظمة التعليمية.'
));

children.push(H2('5. جودة التوصيات'));
children.push(P(
  'جميع النماذج قدمت توصيات مفيدة تركز على بناء الجمل وتوسيع المفردات، وهي توصيات منطقية لمستوى A1. ' +
  'كانت توصيات Mistral Large الأكثر تفصيلاً وعملية، حيث اقترحت "التمرن على المحادثات القصيرة".'
));

// === Recommendations ===
children.push(H1('التوصيات والتطوير المستقبلي'));

children.push(H2('للنظام نفسه'));
children.push(bulletItem('استخدام طبقة تحقق (Voting Layer) لاتخاذ القرار النهائي بناءً على إجماع النماذج بدلاً من نموذج واحد.'));
children.push(bulletItem('تحسين دقة التفريغ الصوتي باستخدام Whisper Large بدل Whisper Small.'));
children.push(bulletItem('إضافة مكون تحليل النطق على مستوى الصوتيات (Phonemes) للتقييم الدقيق للنطق.'));
children.push(bulletItem('بناء قاعدة بيانات مرجعية للأخطاء الشائعة في العربية لزيادة دقة التقييم.'));

children.push(H2('للنماذج'));
children.push(bulletItem('اختبار نماذج متخصصة في العربية مثل Jais و AceGPT لمقارنة أدائها مع النماذج متعددة اللغات.'));
children.push(bulletItem('استخدام Few-shot Prompting بأمثلة تقييم مرجعية لتقليل التباين بين النماذج.'));
children.push(bulletItem('إضافة درجة "ثقة" لكل تقييم لإبراز الحالات الغامضة.'));

children.push(H2('للمستخدم النهائي'));
children.push(bulletItem('إضافة تتبع تاريخي لتقدم المتعلم عبر الزمن.'));
children.push(bulletItem('إضافة تمارين تفاعلية مقترحة بناءً على نقاط الضعف المكتشفة.'));
children.push(bulletItem('دعم اللهجات العربية المختلفة (خليجية، شامية، مصرية، مغاربية).'));

// === Conclusion ===
children.push(H1('الخاتمة'));
children.push(P(
  'أثبت هذا المشروع جدوى استخدام نماذج LLM المتاحة عبر OpenRouter لبناء نظام تقييم آلي للعربية المنطوقة. ' +
  'النتائج تشير إلى أن النماذج الحالية قادرة على إعطاء تقييم مبدئي معقول، خاصةً عند تحديد المستوى العام وفق CEFR. ' +
  'ومع ذلك، فإن التباين الكبير في الدرجات الرقمية يستدعي تطوير منهجية أكثر صرامة للحصول على نتائج موحدة وقابلة للاعتماد.'
));
children.push(P(
  'تشكل هذه النسخة الأولى من المشروع أساساً قوياً يمكن البناء عليه لتطوير منصة تعليمية ذكية كاملة للغة العربية، ' +
  'تخدم ملايين المتعلمين حول العالم الذين يبحثون عن أدوات حديثة لتعلم العربية.'
));

children.push(new Paragraph({
  bidirectional: true, alignment: AlignmentType.CENTER,
  spacing: { before: 600, after: 200 },
  children: [AR('— نهاية التقرير —', { italics: true, color: '8FAADC', size: 24 })],
}));

// =========================================================
// Document setup
// =========================================================
const doc = new Document({
  creator: 'Mujahid',
  title: 'Arabic Speech Assessment Report',
  description: 'تقرير مشروع تقييم النطق العربي بالذكاء الاصطناعي',
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 24 } },
    },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: '1F3864' },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: '2E75B6' },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '•',
        alignment: AlignmentType.RIGHT,
        style: { paragraph: { indent: { right: 360 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
      bidi: true,
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          bidirectional: true,
          alignment: AlignmentType.RIGHT,
          children: [AR('تقرير تقييم النطق العربي', { size: 18, color: '8FAADC' })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ children: ['Page ', PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], size: 18, color: '8FAADC' }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync('Arabic_Speech_Assessment_Report.docx', buf);
  console.log('✓ Created Arabic_Speech_Assessment_Report.docx');
});
