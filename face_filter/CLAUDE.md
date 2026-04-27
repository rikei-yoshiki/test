# face_filter プロジェクト

顔認識でフォルダ内の写真ごとに「誰が映っているか」をカウントしExcel出力するスクリプト。

## フォルダ構成

```
C:\GitHub\test\face_filter\
├── face_count.py             メインスクリプト
├── install_insightface.bat   insightface再インストール用（vcvarsall.bat経由でMSVC有効化）
├── requirements.txt          pip依存パッケージ一覧（環境再現用）
├── face_env\                 Python仮想環境（venv、git管理外）
├── face_database\            顔と名前の対応データベース（個人情報のためgit管理外）
├── output\                   Excel出力先（git管理外）
├── output_debug\             デバッグ用：bbox+名前を描画した画像出力先（git管理外）
└── CLAUDE.md                 このファイル
```

## Python環境

- **venv**: `face_env`（標準venv、Python 3.13.5）
- **有効化**: `face_env\Scripts\activate`（cmd/PowerShell）または `source face_env/Scripts/activate`（bash）

### 主要な依存パッケージ

```
insightface 0.7.3       顔検出+認識（buffalo_lモデル使用、初回DL ~280MB）
onnxruntime 1.25.0      ONNX推論エンジン
opencv-python-headless  画像処理
openpyxl 3.1.5          Excel出力
pillow 12.2.0           画像読み込み・PIL描画
pillow-heif 1.3.0       HEIC対応
numpy 2.4.4
```

完全な依存リストは `requirements.txt` 参照（pip freezeで自動生成、約50パッケージ）。

### ビルド要件

InsightFace は Cython 拡張のビルドに **Visual Studio Build Tools 2022（C++）** が必要。
最小構成（MSVC + Windows 11 SDK のみ、約3-4GB）でOK。

`pip install insightface` を直接叩くとMSVCを見つけられず失敗するため、
`install_insightface.bat`（vcvarsall.bat を読み込んでから pip install するbat）を経由する必要がある。

### 別PCで環境を再現する手順

venv（`face_env/`）はOS・CPUアーキ・Pythonバージョンに依存するためgit管理せず、`requirements.txt` から再構築する。

```bash
# 1. リポジトリをクローン
git clone <repo_url>
cd test/face_filter

# 2. Visual Studio Build Tools 2022（C++）を事前インストール
#    最小構成: VC.Tools.x86.x64 + Windows11SDK.22621
#    winget install --id Microsoft.VisualStudio.2022.BuildTools --override "--passive --add Microsoft.VisualStudio.Component.VC.Tools.x86.x64 --add Microsoft.VisualStudio.Component.Windows11SDK.22621"

# 3. venv作成（Python 3.13系を使用）
python -m venv face_env

# 4. insightfaceだけ先にbat経由でビルド（vcvarsall.bat必須）
install_insightface.bat

# 5. 残りの依存パッケージを一括インストール
face_env\Scripts\python.exe -m pip install -r requirements.txt

# 6. face_databaseフォルダを作成し、顔DB画像を配置
#    ファイル名規則: <クラス名>_<名前>.png または <クラス名>_<名前>_<補助>.png
```

## face_database のファイル名規則

```
<クラス名>_<名前>.png                    例: あか組_竹内未姫.png
<クラス名>_<名前>_<補助>.jpeg            例: あか組_竹内未姫_2.jpeg
```

- 最初の `_` までが クラス名、2番目の `_` までが 名前
- 同じ「クラス名_名前」のファイルは**同一人物**として扱われ、特徴量が平均化される
- 1人につき複数枚登録すると認識精度が上がる（特に検出が難しい顔の場合）
- **重要**：補助番号は必ずアンダースコアで区切ること。`咲音2.png` のようにアンダースコア無しで番号を付けると別人扱いになる

### 現在の登録メンバー（2026-04-28時点 9名）

```
あか組_竹内咲音       (2枚: _2, _3)
あか組_竹内未姫       (3枚: 単体, _2, _3) ※赤ちゃん、検出が難しい
さくら組_竹内美香     (1枚)
しろ組_久冨美都       (2枚: 単体, _2)
すみれ組_久冨賀江     (1枚)
とら組_竹内海愛       (1枚)
みどり組_久冨悠人     (1枚)
昭和組_久冨秀樹       (1枚)
昭和組_竹内雪子       (1枚)
```

### 既知の挙動

- **赤ちゃんの顔は検出されにくい**：成人顔向けに学習されているため、デフォルトしきい値0.5では検出失敗することがある。`build_face_database` 内でしきい値を 0.5→0.3→0.1 と段階的に下げて救済する仕組みがある（DB画像は信頼できる前提なので緩めてOK）
- **EXIF Orientation を考慮**：スマホ撮影写真は横向き保存されていることが多いため、`load_image` で `ImageOps.exif_transpose` を呼んで補正している
- **透過PNGは白背景に合成**：透明部分が0埋めされると検出器が誤動作するため

## 実行方法

```bash
cd C:\GitHub\test\face_filter
face_env\Scripts\python.exe face_count.py
```

## パラメータ（face_count.py 冒頭）

| 名前 | 現在値 | 意味 |
|---|---|---|
| `TARGET_FOLDER` | `\\IESSANsNAS\nas_share\Photo\2025\2503` | 解析対象フォルダ |
| `FACE_DB_FOLDER` | `C:\GitHub\test\face_filter\face_database` | 顔DB |
| `OUTPUT_FOLDER` | `C:\GitHub\test\face_filter\output` | Excel出力先 |
| `OUTPUT_DEBUG_FOLDER` | `C:\GitHub\test\face_filter\output_debug` | デバッグ画像出力先 |
| `DEBUG_DRAW_ANNOTATIONS` | `True` | bbox+名前を描画したデバッグ画像を出力するか |
| `DEBUG_FONT_PATH` | `C:\Windows\Fonts\meiryo.ttc` | 日本語フォントパス |
| `SIMILARITY_THRESHOLD` | `0.60` | 同一人物判定のコサイン類似度しきい値。**取りこぼし多→下げる、誤検出多→上げる** |
| `DETECTION_MAX_SIZE` | `1280` | 検出時の画像最大サイズ（長辺ピクセル） |
| `DETECTION_INPUT_SIZE` | `(640, 640)` | 検出器の入力サイズ |
| `DB_DETECTION_FALLBACK_THRESHOLDS` | `[0.5, 0.3, 0.1]` | DB登録時の段階的しきい値 |
| `RECURSIVE` | `False` | サブフォルダも処理するか |
| `MAX_FILES` | `20` | 処理する最大ファイル数（テスト用、`None`で全件） |
| `EXTENSIONS` | `['.jpg','.jpeg','.png','.heic','.bmp','.webp']` | 対象拡張子 |
| `MODEL_NAME` | `'buffalo_l'` | InsightFaceモデル（`buffalo_s`は軽量版） |

### しきい値調整の経緯

- 当初 0.40 → 誤判定多発（DB未登録の親戚を家族と判定）
- 0.50 推奨 → 0.49程度の誤判定を弾けない
- **0.60 採用**：誤判定明確に減るが、家族でも0.58〜0.59になるケースがあり取りこぼしリスクあり

## Excel出力レイアウト

- 縦軸: ファイル名
- 横軸: 人物名（クラス名_名前）
- マッチは `○`
- 最右列: ファイル別の合計人数（COUNTIF関数）
- 最終行: 人物別の出現ファイル数（COUNTIF関数）
- B2でフリーズ、ヘッダー縦書き

## デバッグ画像出力

`DEBUG_DRAW_ANNOTATIONS = True` のとき、`output_debug/` に注釈付き画像を保存する。

- **緑枠**: マッチ成功 → `クラス名_名前 (類似度)` を表示
- **赤枠**: 顔検出はしたがマッチしなかった顔 → `? (類似度)` を表示
- フォントサイズ・線幅は**顔のbbox高さに比例**（画像全体ではなく顔基準のほうが視認性が安定）

Excelとファイル名を照合するのは大変なので、デバッグ画像で目視確認するのが効率的。

## 設計の意図

### 認識エンジンの抽象化

`FaceRecognizer` 抽象基底クラス + `InsightFaceRecognizer` 実装の構成にしてある。
後で `facenet-pytorch` などへの差し替えが容易（精度や速度の都合で乗り換える場合）。

InsightFace採用の理由：
- 集合写真・小さい顔・横顔に強い
- 業界トップクラスの精度（LFW 99.83%）
- ただしビルド要件が厳しい（VS Build Tools必須）

### マッチング戦略（現状=平均化、将来検討=最大類似度）

現状：同一人物の複数枚DB画像から得た特徴量を**平均化**して1ベクトルにする。

**問題**：子供の0歳と3歳など顔変化が大きい場合、平均すると「どっちでもない中間ベクトル」になり、両方の年齢の写真とも類似度が中途半端になる。

**改善案**：「最大類似度マッチング」に変更する。
- 各サンプルとの類似度を個別に計算
- 最も高いものを採用
- DB写真のバリエーションがそのまま認識力アップに繋がる

実装は `recognize` 関数で `db.items()` をイテレートしている部分を、ラベル単位ではなく「ラベル＋サンプル」単位で比較するよう拡張すればOK。

## DB写真の集め方ガイド

精度向上のためのベストプラクティス：

| サンプル数 | 効果 |
|---|---|
| 1枚 | 単一表情・角度に依存。精度低い |
| 3〜5枚 | 大幅改善 |
| 5〜10枚 | 緩やかに改善 |
| 10枚以上 | ほぼ頭打ち（似た写真ばかりだと無意味） |

### 質 > 量

特に**子どもの写真は時系列で顔が変化**する。多様性を意識する：
- ✅ 異なる年齢／成長段階（特に乳幼児は重要）
- ✅ 異なる角度（正面・斜め・横顔）
- ✅ 異なる表情（真顔・笑顔・口開）
- ✅ 異なる照明（屋内・屋外・夜）
- ❌ 同じ日に連写した10枚 → ほぼ1枚と同じ

## テスト履歴

| 日付 | 件数 | しきい値 | 備考 |
|---|---|---|---|
| 2026-04-27 | 10 | 0.40 | 初回テスト。誤判定発生（0.49で別人を未姫と判定） |
| 2026-04-27 | 10 | 0.60 | 誤判定除外、ただし他の判定数も減少 |
| 2026-04-27 | 20 | 0.60 | DB拡充（咲音, 美都2追加）後。一致数増加を確認 |

未実施：
- C: NAS全253件で本番実行
