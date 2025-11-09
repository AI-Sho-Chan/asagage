# AWS HPC 実行基盤（週末フル探索向け）

目的: 週末の重いフル探索を AWS 上で短時間・低コストに実行し、成果物をローカルへ取り込む。夜間や軽量微調整はローカル継続。

このディレクトリは「設定だけ」を先行提供します（本番プロビジョニングは未実行）。

## 概要
- 方式: Ray Autoscaler + EC2 Spot（c7i 系などの高クロック CPU）
- 構成:
  - ヘッド: 小型オンデマンド（例: t3.large）
  - ワーカー: Spot c7i.8xlarge〜c7i.48xlarge（台数は `max_workers`）
  - AMI: Amazon Linux 2（Ray 推奨）
  - ストレージ: S3 を minute cache / 成果物の中継に利用（任意）

## 必要準備
1) AWS アカウント/プロファイル（`aws configure` 済）
2) キーペア（`ec2-keypair`）
3) S3 バケット（任意: `asagake-cache-<account>`）

## 使い方（雛形）
1. Ray 環境をローカルにインストール
```
pip install "ray[default]"
```

2. クラスタを起動（ドライラン推奨）
```
ray up -y infra/aws/ray-cluster.yaml   # 初回は -y を外して内容確認
```

3. リポジトリをヘッドへ同期
```
ray rsync_up infra/aws/ray-cluster.yaml . /home/ec2-user/asagake
```

4. ヘッドへ接続して実行
```
ray attach infra/aws/ray-cluster.yaml
cd asagake
python -m pip install -r requirements.txt || true
python scripts/nightly_build_candidates.py \
  --excel ASAGAKE.xlsm \
  --base-out output/bt30/WEEKLY_AWS \
  --run-type weekend --plan-profile weekend \
  --lookback 60 --chunk-days 5 --train-days 12 --forward-days 4 \
  --min-train-trades 12 --min-forward-trades 5 --forward-pf-min 1.3 \
  --gap-guard-abs-bp 80 --gap-guard-dir-bp 40 \
  --liquidity-quantile 0.3 --jobs 48 \
  --enable-asha --enable-bayes --bayes-trials 20 --bayes-timeout 600 \
  --mask-ineffective --mask-window 20 --mask-threshold 1.05 \
  --enable-market-features --excel-summary --analysis-ledger
```

5. 成果物の取得
```
ray rsync_down infra/aws/ray-cluster.yaml /home/ec2-user/asagake/output ./output
```

6. 終了
```
ray down -y infra/aws/ray-cluster.yaml
```

## コストと注意
- Spot は大幅に安価ですが、奪還リスクあり（Ray が再スケジュール）。
- まずは `max_workers` を小さめ（2〜4）で検証し、所要時間/費用を確認してから拡張してください。

