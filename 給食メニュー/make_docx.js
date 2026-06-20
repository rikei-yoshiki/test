const DOCX_PATH = 'C:/Users/kunyo/AppData/Roaming/npm/node_modules/docx';
const imgData = require('fs').readFileSync('C:/GitHub/test/給食メニュー/tmp_unpack/extracted/word/media/image1.jpeg');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, HeightRule, ImageRun
} = require(DOCX_PATH);
const fs = require('fs');

const PAGE_W = 11906;
const MARGIN = 850;
const CW = PAGE_W - MARGIN * 2; // 10206

const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const LINE = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const THIN = { style: BorderStyle.SINGLE, size: 2, color: "000000" };

const noBorder  = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const allBorder = { top: LINE, bottom: LINE, left: LINE, right: LINE };
const allThin   = { top: THIN, bottom: THIN, left: THIN, right: THIN };

function p(text, opts = {}) {
  const { bold=false, size=20, align=AlignmentType.LEFT, spacing={}, color="000000" } = opts;
  return new Paragraph({
    alignment: align,
    spacing,
    children: [new TextRun({ text, bold, size, font: "游ゴシック", color })]
  });
}

function tc(content, width, opts = {}) {
  const { borders=allThin, colspan=1, rowspan=1, vAlign=VerticalAlign.CENTER } = opts;
  const children = Array.isArray(content) ? content : (typeof content === 'string' ? [p(content)] : [content]);
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    columnSpan: colspan,
    rowSpan: rowspan,
    verticalAlign: vAlign,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children
  });
}

// 列幅
const W_MAT = 2700;
const W_AMT = 1600;
const W_REC = CW - W_MAT - W_AMT; // 5906

// 材料
const ingredients = [
  ["マカロニ（ゆでたもの）", "50g"],
  ["とり肉（細切れ）", "30g"],
  ["ブロッコリー", "小3房（25g）"],
  ["カリフラワー", "小3房（25g）"],
  ["玉ねぎ", "1/4個（40g）"],
  ["豆乳（無調整）", "100ml"],
  ["薄力粉", "小さじ2"],
  ["バター", "5g"],
  ["塩・こしょう", "少々"],
  ["ピザ用チーズ", "15g"],
  ["", ""],
  ["", ""],
];
const TOTAL_ROWS = ingredients.length + 1;

// 作り方
const steps = [
  "① ブロッコリーとカリフラワーは小さな房に分けて、塩ゆでする。マカロニも塩ゆでしておく。",
  "② 玉ねぎは薄切り、とり肉は一口大に切る。",
  "③ フライパンにバターをとかし、玉ねぎととり肉をいためる。火が通ったら薄力粉をふりかけてよくまぜる。",
  "④ 豆乳を少しずつ加えながら、とろみがつくまでまぜる。塩・こしょうで味をととのえる。",
  "⑤ 耐熱皿にマカロニ、ブロッコリー、カリフラワーを入れて、④のソースをかける。",
  "⑥ チーズをのせてオーブントースターで5分、チーズがとけるまで焼いたらできあがり。",
];

function makeIngredientTable() {
  const rows = [];

  // ヘッダー行（作り方セルはrowspan）
  rows.push(new TableRow({
    children: [
      tc(p("材　料　名", { bold: true, size: 20, align: AlignmentType.CENTER }), W_MAT),
      tc(p("1人分量", { bold: true, size: 20, align: AlignmentType.CENTER }), W_AMT),
      new TableCell({
        width: { size: W_REC, type: WidthType.DXA },
        borders: allThin,
        rowSpan: TOTAL_ROWS,
        verticalAlign: VerticalAlign.TOP,
        margins: { top: 80, bottom: 80, left: 140, right: 100 },
        children: [
          p("作り方（できるだけくわしく書いてください。）", { bold: true, size: 20 }),
          p("", { size: 12, spacing: { after: 60 } }),
          ...steps.map(s => p(s, { size: 19, spacing: { after: 160 } })),
        ]
      })
    ]
  }));

  // 材料行
  for (const [name, amount] of ingredients) {
    rows.push(new TableRow({
      height: { value: 420, rule: HeightRule.ATLEAST },
      children: [
        tc(p(name, { size: 18 }), W_MAT),
        tc(p(amount, { size: 18, align: AlignmentType.CENTER }), W_AMT),
      ]
    }));
  }

  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [W_MAT, W_AMT, W_REC],
    rows
  });
}

// おすすめポイント + 盛り付け図
const W_POINT = Math.floor(CW * 0.44); // 4490
const W_MORI  = CW - W_POINT;          // 5716

const pointLines = [
  "ブロッコリーとカリフラワーは形がそっくりなのに色がちがいます。",
  "いっしょに入れると見た目がにぎやかになります。",
  "",
  "牛乳のかわりに愛知県産の大豆から作った豆乳を使ったので、",
  "ソースがあっさりしていて野菜のあじがよくわかります。",
  "家族に出したら「牛乳じゃないの？」とおどろかれました。",
];

function makeBottomTable() {
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [W_POINT, W_MORI],
    rows: [new TableRow({
      height: { value: 2800, rule: HeightRule.ATLEAST },
      children: [
        new TableCell({
          width: { size: W_POINT, type: WidthType.DXA },
          borders: allThin,
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 80, bottom: 80, left: 140, right: 100 },
          children: [
            p("おすすめポイント", { bold: true, size: 20 }),
            p("", { size: 12, spacing: { after: 60 } }),
            ...pointLines.map(t => p(t, { size: 18, spacing: { after: 40 } })),
          ]
        }),
        new TableCell({
          width: { size: W_MORI, type: WidthType.DXA },
          borders: allThin,
          verticalAlign: VerticalAlign.TOP,
          margins: { top: 80, bottom: 80, left: 140, right: 100 },
          children: [
            p("盛り付け図", { bold: true, size: 20 }),
            p("", { size: 12, spacing: { after: 80 } }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new ImageRun({
                type: "jpg",
                data: imgData,
                transformation: { width: 260, height: 180 },
                altText: { title: "盛り付け図", description: "カリブログラタンのイラスト", name: "盛り付け図" }
              })]
            }),
          ]
        }),
      ]
    })]
  });
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    children: [
      // タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "「愛知の産物を使った我が家のじまん料理」", bold: true, size: 34, font: "游ゴシック" })]
      }),

      // 注意（右寄せ）
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 100 },
        children: [new TextRun({ text: "※わすれずにふりがなを記入しましょう。", size: 16, font: "游ゴシック", color: "555555" })]
      }),

      // 学校・氏名テーブル
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [1800, 2600, 1100, 900, 1100, 2706],
        rows: [
          // ふりがな行
          new TableRow({
            height: { value: 260, rule: HeightRule.EXACT },
            children: [
              tc(p("ふりがな", { size: 14, color: "666666" }), 1800, { borders: { top: LINE, bottom: NONE, left: LINE, right: THIN } }),
              tc(p("", { size: 14 }), 2600, { borders: { top: LINE, bottom: NONE, left: THIN, right: LINE } }),
              tc(p("", { size: 14 }), 1100, { borders: { top: LINE, bottom: NONE, left: LINE, right: THIN } }),
              tc(p("", { size: 14 }), 900,  { borders: { top: LINE, bottom: NONE, left: THIN, right: LINE } }),
              tc(p("ふりがな", { size: 14, color: "666666" }), 1100, { borders: { top: LINE, bottom: NONE, left: LINE, right: THIN } }),
              tc(p("", { size: 14 }), 2706, { borders: { top: LINE, bottom: NONE, left: THIN, right: LINE } }),
            ]
          }),
          // 学校・氏名行
          new TableRow({
            height: { value: 500, rule: HeightRule.ATLEAST },
            children: [
              new TableCell({
                width: { size: 1800, type: WidthType.DXA },
                borders: allThin,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "　　　　　　　", size: 22, font: "游ゴシック" }), new TextRun({ text: "小学校", bold: true, size: 22, font: "游ゴシック" })] })]
              }),
              new TableCell({
                width: { size: 2600, type: WidthType.DXA },
                borders: allThin,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "", size: 22, font: "游ゴシック" })] })]
              }),
              new TableCell({
                width: { size: 1100, type: WidthType.DXA },
                borders: allThin,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "　　年", bold: true, size: 22, font: "游ゴシック" })] })]
              }),
              new TableCell({
                width: { size: 900, type: WidthType.DXA },
                borders: allThin,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "　　組", bold: true, size: 22, font: "游ゴシック" })] })]
              }),
              tc(p("氏　名", { bold: true, size: 22, align: AlignmentType.CENTER }), 1100),
              tc(p("", { size: 22 }), 2706),
            ]
          }),
        ]
      }),

      // 料理名
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [1600, CW - 1600],
        rows: [new TableRow({
          height: { value: 620, rule: HeightRule.ATLEAST },
          children: [
            tc(p("料　理　名", { bold: true, size: 22, align: AlignmentType.CENTER }), 1600),
            new TableCell({
              width: { size: CW - 1600, type: WidthType.DXA },
              borders: allThin,
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 60, bottom: 60, left: 150, right: 100 },
              children: [
                p("カリブロ大作戦！愛知豆乳グラタン", { bold: true, size: 28 }),
                p("（かりぶろ だいさくせん！あいち とうにゅう ぐらたん）", { size: 15, color: "555555" }),
              ]
            }),
          ]
        })]
      }),

      // 愛知の産物
      new Table({
        width: { size: CW, type: WidthType.DXA },
        columnWidths: [CW],
        rows: [new TableRow({
          height: { value: 460, rule: HeightRule.ATLEAST },
          children: [
            new TableCell({
              width: { size: CW, type: WidthType.DXA },
              borders: allThin,
              margins: { top: 60, bottom: 60, left: 140, right: 100 },
              children: [new Paragraph({
                children: [
                  new TextRun({ text: "愛知の産物（", bold: true, size: 22, font: "游ゴシック" }),
                  new TextRun({ text: "大豆（豆乳）、ブロッコリー、カリフラワー、玉ねぎ、とり肉", size: 22, font: "游ゴシック" }),
                  new TextRun({ text: "）", bold: true, size: 22, font: "游ゴシック" }),
                ]
              })]
            }),
          ]
        })]
      }),

      // 材料・作り方テーブル
      makeIngredientTable(),

      // おすすめポイント + 盛り付け図
      makeBottomTable(),

      // 末尾注意書き
      new Paragraph({
        spacing: { before: 120, after: 0 },
        children: [new TextRun({ text: "※入選作品は、レシピ集に掲載されます。ていねいな字で太く濃く書きましょう。（文字は、できるだけ2B以上のえんぴつか黒色ペンで書きましょう。）", size: 16, font: "游ゴシック", color: "444444" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('C:/GitHub/test/給食メニュー/カリブロ大作戦_愛知豆乳グラタン_v3.docx', buf);
  console.log('done');
}).catch(e => { console.error(e.message); process.exit(1); });
