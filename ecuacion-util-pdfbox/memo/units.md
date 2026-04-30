# 単位メモ：pt / px / DPI

## pt（ポイント）とは

**1pt = 1/72 インチ** という物理的な長さの単位。  
PDF も Excel も同じ定義で使っているので、両者の間で変換不要。

- Excel の行高・マージン → pt 指定
- PDF の座標系 → pt ベース
- 例：Excel で行高 60pt → PDF でも 60pt（そのまま対応）

## px（ピクセル）とは

DPI（画面解像度）に依存する単位。「1インチに何ドット並ぶか」で決まる。

```
1 px = 1/DPI インチ
1 pt = (DPI/72) px
```

## DPI と pt→px 変換の早見表

| DPI | 1pt = ? px | 用途 |
|-----|----------:|------|
| 72  | 1 px      | PDFBox デフォルト、pt と 1:1 で楽 |
| 96  | 1.33 px   | Windows スクリーン標準（POI の内部計算で登場） |
| 144 | 2 px      | テスト用途で使う（後述） |

## テストで 144 DPI を使う理由

PDFBox の `PDFRenderer.renderImageWithDPI(page, dpi)` でレンダリングした画像を  
ピクセル単位で検査する際、DPI が低いと細い線が半画素になってアンチエイリアスで曖昧になる。

| 線幅 | 72 DPI | 144 DPI |
|------|-------:|--------:|
| THIN（0.5pt）  | 0.5 px → 曖昧 | **1 px** → 明確 |
| THICK（1.5pt） | 1.5 px → 曖昧 | **3 px** → 明確 |

→ 144 DPI にすることで整数ピクセルになり、「暗い／白い」の閾値判定が安定する。

## POI の列幅単位

Excel の列幅は pt でも px でもなく **「最大桁文字幅の 1/256」** という独自単位。

```
POI 単位 → px（96 DPI）:  POI単位 / 256 × maxDigitWidth(≒7px)
px（96 DPI） → pt:        px × (72/96) = px × 0.75
```

例：`sheet.setColumnWidth(0, 2048)`  
→ 2048 / 256 × 7 = 56 px（96 DPI）→ 56 × 0.75 = **42 pt**

## SheetRenderer 内の定数

```java
// SheetRenderer.java
private static final float PX_TO_PT = 72f / 96f; // = 0.75
```

列幅ピクセル（96 DPI）→ pt への変換に使っている。
