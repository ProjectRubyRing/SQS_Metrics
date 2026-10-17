#!/usr/bin/env bash
#==============================================================================
#  sqs_metrics_report.sh
#
#  概要:
#    指定した SQS キュー（および DLQ）の CloudWatch メトリクスを取得し、
#    指定した時間範囲のメッセージ追加状況を「1時間単位」で集計した
#    Excel レポート(.xlsx)を出力する。
#
#    - メインキュー: NumberOfMessagesSent 等から時間別の追加数・受信数・削除数・滞留数を集計
#    - DLQ        : 総メッセージ数(可視+処理中+遅延)の増加分から時間別の追加数を推定
#                   ※ maxReceiveCount 超過による DLQ への自動移動は
#                     DLQ の NumberOfMessagesSent に計上されない(AWS仕様)ため
#
#  動作環境:
#    RHEL 9 系 (bash 5.1 / GNU coreutils / python3 3.9 ※標準ライブラリのみ使用)
#    AWS CLI v2 (v1 でも動作可)
#
#  必要な IAM 権限:
#    sqs:GetQueueUrl, sqs:GetQueueAttributes, cloudwatch:GetMetricData
#
#  使い方は -h を参照。
#==============================================================================
set -euo pipefail

readonly SCRIPT_NAME="${0##*/}"
readonly SCRIPT_VERSION="1.0.0"
readonly DETAIL_LOOKBACK_SEC=3600                 # 期間開始直前の基準値算出用に詳細データを先読みする秒数
readonly BASELINE_LOOKBACK_SEC=$((14 * 86400))    # 基準値を探索する範囲(SQS の最大メッセージ保持期間)
readonly MAX_DETAIL_ROWS=1000000                  # 詳細シートの最大行数(Excel の行数上限対策)

# 実行コマンド(環境変数で差し替え可能)
AWS_BIN="${AWS_BIN:-aws}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

# 引数
QUEUE_INPUT=""
DLQ_INPUT=""
START_INPUT=""
END_INPUT=""
TZ_NAME="Asia/Tokyo"
GRANULARITY="auto"
REGION=""
PROFILE=""
OUTPUT=""
KEEP_RAW=0

# 実行時に決定する値
WORK_DIR=""
PY_SCRIPT=""
NOW_EPOCH=""
START_EPOCH=""
END_EPOCH=""
FETCH_START_EPOCH=""
PERIOD=""
MAIN_NAME=""
MAIN_URL=""
DLQ_NAME=""
DLQ_URL=""
DLQ_SOURCE=""
NOTES=()

#------------------------------------------------------------------------------
# 共通関数
#------------------------------------------------------------------------------
usage() {
  cat <<EOF
使い方:
  ${SCRIPT_NAME} -q <キュー名|キューURL> [-s 開始日時] [-e 終了日時] [オプション]

  指定した SQS キューと DLQ の CloudWatch メトリクスから、時間単位のメッセージ
  追加状況を集計した Excel レポート(.xlsx)を出力します。

必須:
  -q, --queue <名前|URL>      対象の SQS キュー名またはキュー URL

オプション:
  -s, --start <日時>          集計開始日時 (既定: 終了日時の24時間前)
  -e, --end <日時>            集計終了日時 (既定: 現在時刻)
                              ※ 開始は時単位で切り捨て、終了は時単位で切り上げます
                              ※ 指定例: "2026-10-01 09:00" / "2026/10/01 09:00" /
                                        202610010900 / "yesterday 00:00" / "3 days ago"
  -d, --dlq <名前|URL|none>   DLQ (省略時は RedrivePolicy から自動検出、none で DLQ 集計なし)
  -z, --timezone <TZ>         日時の解釈・表示に使うタイムゾーン (既定: Asia/Tokyo)
  -g, --granularity <秒>      詳細データの粒度 auto|60|300|3600 (既定: auto)
                              auto: CloudWatch の保持期間に応じて 60秒(15日以内) →
                                    300秒(63日以内) → 3600秒 の順に自動選択
  -r, --region <リージョン>    AWS リージョン (省略時: キュー URL / AWS CLI 設定から決定)
  -p, --profile <名前>         AWS CLI のプロファイル
  -o, --output <パス>         出力ファイル
                              (既定: ./sqs_report_<キュー名>_<開始>-<終了>.xlsx)
  -k, --keep-raw              取得した生データ(JSON)を <出力ファイル名>_raw/ に保存
  -h, --help                  このヘルプを表示

例:
  ${SCRIPT_NAME} -q my-queue -s "2026-10-16 00:00" -e "2026-10-17 00:00"
  ${SCRIPT_NAME} -q my-queue -s "3 days ago" -o /var/tmp/my-queue.xlsx
  ${SCRIPT_NAME} -q https://sqs.ap-northeast-1.amazonaws.com/123456789012/my-queue -d my-dlq

出力シート:
  サマリー / 時間別レポート / グラフ / 時間別_メインキュー / 時間別_DLQ /
  DLQ追加一覧 / 詳細_メインキュー / 詳細_DLQ / 指標の説明

終了コード:
  0: 正常終了  1: 引数エラー  2: 実行環境エラー  3: AWS API エラー  4: レポート生成エラー
EOF
}

log() {
  local level="$1"
  shift
  printf '%s [%-5s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "${level}" "$*" >&2
}

warn() {
  log WARN "$*"
  NOTES+=("$*")
}

die() {
  local code="$1"
  shift
  log ERROR "$*"
  exit "${code}"
}

cleanup() {
  if [[ -n "${WORK_DIR}" && -d "${WORK_DIR}" ]]; then
    rm -rf -- "${WORK_DIR}"
  fi
}

need_arg() {
  if [[ -z "${2:-}" ]]; then
    die 1 "オプション $1 には値を指定してください (-h でヘルプを表示)"
  fi
}

aws_cmd() {
  local -a opts=()
  if [[ -n "${REGION}" ]]; then
    opts+=(--region "${REGION}")
  fi
  if [[ -n "${PROFILE}" ]]; then
    opts+=(--profile "${PROFILE}")
  fi
  AWS_PAGER="" "${AWS_BIN}" "${opts[@]}" "$@"
}

# AWS CLI を実行し、標準出力をファイルに保存する。失敗時は終了する。
#   $1: 処理名  $2: 出力ファイル  $3以降: aws の引数
aws_to_file() {
  local desc="$1" out="$2"
  shift 2
  if ! aws_cmd "$@" >"${out}" 2>"${WORK_DIR}/aws_err.log"; then
    die 3 "${desc}に失敗しました: $(tr -s '\r\n' '  ' <"${WORK_DIR}/aws_err.log")"
  fi
}

run_py() {
  PYTHONIOENCODING=utf-8 "${PYTHON_BIN}" "${PY_SCRIPT}" "$@"
}

# 日時文字列を epoch 秒に変換する (TZ_NAME で解釈)
to_epoch() {
  local s="$1"
  if [[ "${s}" =~ ^([0-9]{4})([0-9]{2})([0-9]{2})([0-9]{2})([0-9]{2})$ ]]; then
    s="${BASH_REMATCH[1]}-${BASH_REMATCH[2]}-${BASH_REMATCH[3]} ${BASH_REMATCH[4]}:${BASH_REMATCH[5]}"
  elif [[ "${s}" =~ ^[0-9]{4}/[0-9]{1,2}/[0-9]{1,2} ]]; then
    s="${s//\//-}"
  fi
  TZ="${TZ_NAME}" date -d "${s}" +%s 2>/dev/null
}

# epoch 秒を TZ_NAME の時単位に切り捨て / 切り上げる
floor_hour() {
  local s
  s="$(TZ="${TZ_NAME}" date -d "@$1" '+%Y-%m-%d %H:00:00')"
  TZ="${TZ_NAME}" date -d "${s}" +%s
}

ceil_hour() {
  local f
  f="$(floor_hour "$1")"
  if ((f < $1)); then
    f=$((f + 3600))
  fi
  echo "${f}"
}

iso_utc() {
  date -u -d "@$1" '+%Y-%m-%dT%H:%M:%SZ'
}

fmt_local() {
  TZ="${TZ_NAME}" date -d "@$1" "${2:-+%Y-%m-%d %H:%M}"
}

#------------------------------------------------------------------------------
# 引数解析・事前チェック
#------------------------------------------------------------------------------
parse_args() {
  while (($# > 0)); do
    # --opt=value 形式を --opt value に分解
    if [[ "$1" == --*=* ]]; then
      set -- "${1%%=*}" "${1#*=}" "${@:2}"
    fi
    case "$1" in
      -q | --queue)       need_arg "$1" "${2:-}"; QUEUE_INPUT="$2"; shift 2 ;;
      -d | --dlq)         need_arg "$1" "${2:-}"; DLQ_INPUT="$2"; shift 2 ;;
      -s | --start)       need_arg "$1" "${2:-}"; START_INPUT="$2"; shift 2 ;;
      -e | --end)         need_arg "$1" "${2:-}"; END_INPUT="$2"; shift 2 ;;
      -z | --timezone)    need_arg "$1" "${2:-}"; TZ_NAME="$2"; shift 2 ;;
      -g | --granularity) need_arg "$1" "${2:-}"; GRANULARITY="$2"; shift 2 ;;
      -r | --region)      need_arg "$1" "${2:-}"; REGION="$2"; shift 2 ;;
      -p | --profile)     need_arg "$1" "${2:-}"; PROFILE="$2"; shift 2 ;;
      -o | --output)      need_arg "$1" "${2:-}"; OUTPUT="$2"; shift 2 ;;
      -k | --keep-raw)    KEEP_RAW=1; shift ;;
      -h | --help)        usage; exit 0 ;;
      *)                  usage >&2; die 1 "不明な引数です: $1" ;;
    esac
  done

  if [[ -z "${QUEUE_INPUT}" ]]; then
    usage >&2
    die 1 "対象キュー(-q)を指定してください"
  fi
  case "${GRANULARITY}" in
    auto | 60 | 300 | 3600) ;;
    *) die 1 "粒度(-g)は auto / 60 / 300 / 3600 のいずれかを指定してください: ${GRANULARITY}" ;;
  esac
}

check_prereqs() {
  command -v "${AWS_BIN}" >/dev/null 2>&1 ||
    die 2 "AWS CLI が見つかりません (${AWS_BIN})。AWS CLI v2 をインストールしてください"
  command -v "${PYTHON_BIN}" >/dev/null 2>&1 ||
    die 2 "python3 が見つかりません。'dnf install python3' でインストールしてください"
  "${PYTHON_BIN}" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 6) else 1)' ||
    die 2 "Python 3.6 以上が必要です"
  date -d '@0' +%s >/dev/null 2>&1 ||
    die 2 "GNU date が必要です"
  if [[ "${TZ_NAME}" != "UTC" && -d /usr/share/zoneinfo && ! -f "/usr/share/zoneinfo/${TZ_NAME}" ]]; then
    die 1 "不明なタイムゾーンです: ${TZ_NAME} (例: Asia/Tokyo, UTC)"
  fi
}

resolve_times() {
  local start_raw end_raw now_ceil fetch_age max_period

  NOW_EPOCH="$(date +%s)"
  if [[ -n "${END_INPUT}" ]]; then
    end_raw="$(to_epoch "${END_INPUT}")" || die 1 "終了日時を解釈できません: ${END_INPUT}"
  else
    end_raw="${NOW_EPOCH}"
  fi
  if [[ -n "${START_INPUT}" ]]; then
    start_raw="$(to_epoch "${START_INPUT}")" || die 1 "開始日時を解釈できません: ${START_INPUT}"
  else
    start_raw=$(($(ceil_hour "${end_raw}") - 86400))
  fi

  START_EPOCH="$(floor_hour "${start_raw}")"
  END_EPOCH="$(ceil_hour "${end_raw}")"
  now_ceil="$(ceil_hour "${NOW_EPOCH}")"
  if ((END_EPOCH > now_ceil)); then
    warn "終了日時が未来のため $(fmt_local "${now_ceil}") に補正しました"
    END_EPOCH="${now_ceil}"
  fi
  if ((END_EPOCH <= START_EPOCH)); then
    die 1 "終了日時は開始日時より後を指定してください (開始: $(fmt_local "${START_EPOCH}") / 終了: $(fmt_local "${END_EPOCH}"))"
  fi
  if ((END_EPOCH > NOW_EPOCH)); then
    warn "集計期間に現在時刻を含むため、最後の時間帯($(fmt_local $((END_EPOCH - 3600)) '+%H:00')台)は集計途中の値です"
  fi
  if ((NOW_EPOCH - START_EPOCH > 455 * 86400)); then
    warn "CloudWatch のデータ保持期間(455日)より前の時間帯はデータが存在しません"
  fi

  # 詳細データの粒度: CloudWatch の保持期間(60秒=15日 / 300秒=63日 / 3600秒=455日)を考慮
  FETCH_START_EPOCH=$((START_EPOCH - DETAIL_LOOKBACK_SEC))
  fetch_age=$((NOW_EPOCH - FETCH_START_EPOCH))
  if ((fetch_age <= 15 * 86400 - 3600)); then
    max_period=60
  elif ((fetch_age <= 63 * 86400 - 3600)); then
    max_period=300
  else
    max_period=3600
  fi
  if [[ "${GRANULARITY}" == "auto" ]]; then
    PERIOD="${max_period}"
  else
    PERIOD="${GRANULARITY}"
    if ((PERIOD < max_period)); then
      warn "開始日時が古いため ${PERIOD}秒粒度のデータは CloudWatch に保持されていない可能性があります (推奨: ${max_period}秒以上)"
    fi
  fi
  if (((END_EPOCH - START_EPOCH) / PERIOD > MAX_DETAIL_ROWS)); then
    die 1 "詳細データの行数が上限(${MAX_DETAIL_ROWS})を超えます。期間を短くするか粒度(-g)を大きくしてください"
  fi
}

resolve_output() {
  local base dir
  base="${QUEUE_INPUT%/}"
  base="${base##*/}"
  if [[ -z "${OUTPUT}" ]]; then
    OUTPUT="./sqs_report_${base//[^A-Za-z0-9._-]/_}_$(fmt_local "${START_EPOCH}" +%Y%m%d%H%M)-$(fmt_local "${END_EPOCH}" +%Y%m%d%H%M).xlsx"
  fi
  dir="$(dirname -- "${OUTPUT}")"
  [[ -d "${dir}" && -w "${dir}" ]] || die 1 "出力先ディレクトリが存在しないか書き込みできません: ${dir}"
}

#------------------------------------------------------------------------------
# AWS からの情報取得
#------------------------------------------------------------------------------
region_from_url() {
  if [[ "$1" =~ ^https?://sqs\.([a-z0-9-]+)\.amazonaws\.com ]] ||
    [[ "$1" =~ ^https?://([a-z0-9-]+)\.queue\.amazonaws\.com ]]; then
    echo "${BASH_REMATCH[1]}"
  fi
}

resolve_queues() {
  local arn_region

  # メインキュー
  if [[ "${QUEUE_INPUT}" =~ ^https?:// ]]; then
    MAIN_URL="${QUEUE_INPUT%/}"
    MAIN_NAME="${MAIN_URL##*/}"
    if [[ -z "${REGION}" ]]; then
      REGION="$(region_from_url "${MAIN_URL}")"
    fi
  else
    MAIN_NAME="${QUEUE_INPUT}"
    aws_to_file "キュー URL の取得(${MAIN_NAME})" "${WORK_DIR}/main_url.txt" \
      sqs get-queue-url --queue-name "${MAIN_NAME}" --query QueueUrl --output text
    MAIN_URL="$(tr -d '\r\n' <"${WORK_DIR}/main_url.txt")"
  fi
  log INFO "メインキュー: ${MAIN_NAME} (${MAIN_URL})"
  aws_to_file "キュー属性の取得(${MAIN_NAME})" "${WORK_DIR}/main_attrs.json" \
    sqs get-queue-attributes --queue-url "${MAIN_URL}" --attribute-names All --output json

  # CloudWatch はキューと同じリージョンで参照する
  if [[ -z "${REGION}" ]]; then
    arn_region="$(run_py attr --attrs "${WORK_DIR}/main_attrs.json" --name QueueArn | cut -d: -f4)"
    REGION="${arn_region}"
  fi

  # DLQ
  case "${DLQ_INPUT}" in
    "")
      DLQ_NAME="$(run_py dlq-name --attrs "${WORK_DIR}/main_attrs.json")"
      if [[ -n "${DLQ_NAME}" ]]; then
        DLQ_SOURCE="auto"
      else
        DLQ_SOURCE="notfound"
        warn "メインキューに RedrivePolicy(DLQ) が設定されていないため、DLQ の集計は行いません"
      fi
      ;;
    none | NONE | None)
      DLQ_SOURCE="none"
      ;;
    http://* | https://*)
      DLQ_URL="${DLQ_INPUT%/}"
      DLQ_NAME="${DLQ_URL##*/}"
      DLQ_SOURCE="arg"
      ;;
    *)
      DLQ_NAME="${DLQ_INPUT}"
      DLQ_SOURCE="arg"
      ;;
  esac

  if [[ -n "${DLQ_NAME}" ]]; then
    if [[ -z "${DLQ_URL}" ]]; then
      if aws_cmd sqs get-queue-url --queue-name "${DLQ_NAME}" --query QueueUrl --output text \
        >"${WORK_DIR}/dlq_url.txt" 2>"${WORK_DIR}/aws_err.log"; then
        DLQ_URL="$(tr -d '\r\n' <"${WORK_DIR}/dlq_url.txt")"
      else
        warn "DLQ の URL を取得できませんでした(${DLQ_NAME})。メトリクスのみで集計します: $(tr -s '\r\n' '  ' <"${WORK_DIR}/aws_err.log")"
      fi
    fi
    if [[ -n "${DLQ_URL}" ]]; then
      if ! aws_cmd sqs get-queue-attributes --queue-url "${DLQ_URL}" --attribute-names All --output json \
        >"${WORK_DIR}/dlq_attrs.json" 2>"${WORK_DIR}/aws_err.log"; then
        rm -f "${WORK_DIR}/dlq_attrs.json"
        warn "DLQ の属性を取得できませんでした(${DLQ_NAME}): $(tr -s '\r\n' '  ' <"${WORK_DIR}/aws_err.log")"
      fi
    fi
    log INFO "DLQ: ${DLQ_NAME} (${DLQ_URL:-URL不明}) [$([[ "${DLQ_SOURCE}" == "auto" ]] && echo 'RedrivePolicyから自動検出' || echo '引数指定')]"
  fi
}

fetch_metrics() {
  local -a dlq_opt=()
  if [[ -n "${DLQ_NAME}" ]]; then
    dlq_opt=(--dlq "${DLQ_NAME}")
  fi

  log INFO "集計期間: $(fmt_local "${START_EPOCH}") ～ $(fmt_local "${END_EPOCH}") (${TZ_NAME}) / 詳細粒度: ${PERIOD}秒 / リージョン: ${REGION:-AWS CLI 既定}"

  # 詳細データ(期間開始直前の1時間を含む)
  run_py queries --kind detail --period "${PERIOD}" --main "${MAIN_NAME}" "${dlq_opt[@]}" \
    >"${WORK_DIR}/queries_detail.json" || die 4 "メトリクス取得クエリの生成に失敗しました"
  log INFO "CloudWatch メトリクスを取得しています(詳細データ)..."
  aws_to_file "CloudWatch メトリクス(詳細データ)の取得" "${WORK_DIR}/metrics_detail.json" \
    cloudwatch get-metric-data \
    --metric-data-queries "file://${WORK_DIR}/queries_detail.json" \
    --start-time "$(iso_utc "${FETCH_START_EPOCH}")" \
    --end-time "$(iso_utc "${END_EPOCH}")" \
    --scan-by TimestampAscending --output json

  # 期間開始時点の総メッセージ数を求めるための基準データ(開始前14日間・1時間粒度)
  run_py queries --kind baseline --period 3600 --main "${MAIN_NAME}" "${dlq_opt[@]}" \
    >"${WORK_DIR}/queries_baseline.json" || die 4 "メトリクス取得クエリの生成に失敗しました"
  log INFO "CloudWatch メトリクスを取得しています(期間開始前の基準データ)..."
  aws_to_file "CloudWatch メトリクス(基準データ)の取得" "${WORK_DIR}/metrics_baseline.json" \
    cloudwatch get-metric-data \
    --metric-data-queries "file://${WORK_DIR}/queries_baseline.json" \
    --start-time "$(iso_utc $((START_EPOCH - BASELINE_LOOKBACK_SEC)))" \
    --end-time "$(iso_utc "${START_EPOCH}")" \
    --scan-by TimestampAscending --output json
}

build_report() {
  local note raw_dir
  local -a args=(
    report
    "--output=${OUTPUT}"
    "--main-name=${MAIN_NAME}"
    "--main-url=${MAIN_URL}"
    "--main-attrs=${WORK_DIR}/main_attrs.json"
    "--metrics=${WORK_DIR}/metrics_detail.json"
    "--baseline=${WORK_DIR}/metrics_baseline.json"
    "--start=${START_EPOCH}"
    "--end=${END_EPOCH}"
    "--now=${NOW_EPOCH}"
    "--period=${PERIOD}"
    "--tz=${TZ_NAME}"
    "--tz-offset=$(TZ="${TZ_NAME}" date -d "@${START_EPOCH}" +%z)"
    "--region=${REGION}"
    "--dlq-source=${DLQ_SOURCE}"
    "--host=$(uname -n)"
    "--version=${SCRIPT_VERSION}"
  )
  if [[ -n "${DLQ_NAME}" ]]; then
    args+=("--dlq-name=${DLQ_NAME}" "--dlq-url=${DLQ_URL}")
    if [[ -f "${WORK_DIR}/dlq_attrs.json" ]]; then
      args+=("--dlq-attrs=${WORK_DIR}/dlq_attrs.json")
    fi
  fi
  for note in "${NOTES[@]}"; do
    args+=("--note=${note}")
  done

  log INFO "Excel レポートを作成しています..."
  run_py "${args[@]}" || die 4 "Excel レポートの作成に失敗しました"

  if ((KEEP_RAW)); then
    raw_dir="${OUTPUT%.xlsx}_raw"
    mkdir -p -- "${raw_dir}"
    cp -p -- "${WORK_DIR}"/*.json "${raw_dir}/"
    log INFO "取得した生データを保存しました: ${raw_dir}"
  fi
  log INFO "レポートを出力しました: ${OUTPUT}"
}

#------------------------------------------------------------------------------
# レポート作成処理 (Python 3.6+ / 標準ライブラリのみ)
#------------------------------------------------------------------------------
write_python() {
  PY_SCRIPT="${WORK_DIR}/sqs_report.py"
  cat >"${PY_SCRIPT}" <<'PYEOF'
# -*- coding: utf-8 -*-
"""sqs_metrics_report.sh の内部処理
  queries  : GetMetricData のクエリ(JSON)を生成
  attr     : get-queue-attributes の結果から属性値を出力
  dlq-name : RedrivePolicy から DLQ 名を出力
  report   : 取得したメトリクスを集計して Excel(.xlsx) レポートを作成
"""
import argparse
import calendar
import datetime as dt
import json
import math
import os
import re
import sys
import zipfile
from types import SimpleNamespace
from xml.sax.saxutils import escape

HOUR = 3600

# (キー, CloudWatch メトリクス名, 統計, 集計方法 sum=合計 / max=最大)
METRICS = [
    ('sent',       'NumberOfMessagesSent',                  'Sum',         'sum'),
    ('received',   'NumberOfMessagesReceived',              'Sum',         'sum'),
    ('deleted',    'NumberOfMessagesDeleted',               'Sum',         'sum'),
    ('empty',      'NumberOfEmptyReceives',                 'Sum',         'sum'),
    ('visible',    'ApproximateNumberOfMessagesVisible',    'Maximum',     'max'),
    ('notvisible', 'ApproximateNumberOfMessagesNotVisible', 'Maximum',     'max'),
    ('delayed',    'ApproximateNumberOfMessagesDelayed',    'Maximum',     'max'),
    ('age',        'ApproximateAgeOfOldestMessage',         'Maximum',     'max'),
    ('size_sum',   'SentMessageSize',                       'Sum',         'sum'),
    ('size_cnt',   'SentMessageSize',                       'SampleCount', 'sum'),
    ('size_max',   'SentMessageSize',                       'Maximum',     'max'),
]
METRIC_NAME = {k: m for k, m, _s, _h in METRICS}
# 総メッセージ数 = 可視 + 処理中 + 遅延
STOCK_KEYS = ('visible', 'notvisible', 'delayed')


# =============================================================================
# 補助コマンド
# =============================================================================
def metric_query(qid, queue, metric, stat, period):
    return {
        'Id': qid,
        'MetricStat': {
            'Metric': {
                'Namespace': 'AWS/SQS',
                'MetricName': metric,
                'Dimensions': [{'Name': 'QueueName', 'Value': queue}],
            },
            'Period': period,
            'Stat': stat,
        },
        'ReturnData': True,
    }


def cmd_queries(a):
    targets = [('m', a.main)] + ([('d', a.dlq)] if a.dlq else [])
    queries = []
    for prefix, queue in targets:
        if a.kind == 'detail':
            for key, metric, stat, _how in METRICS:
                queries.append(metric_query(f'{prefix}_{key}', queue, metric, stat, a.period))
        else:
            # 期間開始前の総メッセージ数(1時間ごとの最小値)
            for key in STOCK_KEYS:
                queries.append(metric_query(f'b{prefix}_{key}', queue, METRIC_NAME[key], 'Minimum', HOUR))
    json.dump(queries, sys.stdout, indent=1)


def load_attrs(path):
    if not path or not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        try:
            return json.load(f).get('Attributes') or {}
        except ValueError:
            return {}


def cmd_attr(a):
    print(load_attrs(a.attrs).get(a.name, ''))


def parse_redrive(attrs):
    try:
        return json.loads(attrs.get('RedrivePolicy') or '{}')
    except ValueError:
        return {}


def cmd_dlq_name(a):
    arn = parse_redrive(load_attrs(a.attrs)).get('deadLetterTargetArn', '')
    print(arn.rsplit(':', 1)[-1] if arn else '')


# =============================================================================
# 日時
# =============================================================================
TZ = dt.timezone.utc
EXCEL_EPOCH = dt.datetime(1899, 12, 30)
_TS_RE = re.compile(r'^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})(?:\.\d+)?(Z|[+-]\d{2}:?\d{2})?$')


def set_tz(name, offset):
    global TZ
    try:
        from zoneinfo import ZoneInfo
        TZ = ZoneInfo(name)
        return
    except Exception:
        pass
    # zoneinfo が使えない環境では開始時点の UTC オフセットで固定
    m = re.match(r'^([+-])(\d{2}):?(\d{2})$', offset or '')
    if m:
        delta = dt.timedelta(hours=int(m.group(2)), minutes=int(m.group(3)))
        TZ = dt.timezone(-delta if m.group(1) == '-' else delta, name)


def parse_ts(s):
    """CloudWatch のタイムスタンプ(ISO 8601 / epoch)を epoch 秒に変換"""
    if isinstance(s, (int, float)):
        return int(s)
    s = s.strip()
    if re.match(r'^\d+(\.\d+)?$', s):
        return int(float(s))
    m = _TS_RE.match(s)
    if not m:
        raise ValueError(f'タイムスタンプを解釈できません: {s}')
    epoch = calendar.timegm(tuple(int(m.group(i)) for i in range(1, 7)) + (0, 0, 0))
    tz = m.group(7)
    if tz and tz != 'Z':
        sign = 1 if tz[0] == '+' else -1
        epoch -= sign * (int(tz[1:3]) * 3600 + int(tz[-2:]) * 60)
    return epoch


def local(epoch):
    return dt.datetime.fromtimestamp(epoch, TZ)


def xl_time(epoch):
    """epoch 秒 → Excel のシリアル値(表示タイムゾーンの壁時計時刻)"""
    return (local(epoch).replace(tzinfo=None) - EXCEL_EPOCH).total_seconds() / 86400.0


def fmt_time(epoch, f='%Y/%m/%d %H:%M'):
    return local(epoch).strftime(f)


def hour_label(hs):
    return f'{fmt_time(hs)}～{fmt_time(hs + HOUR, "%H:%M")}'


def fmt_duration(sec):
    try:
        sec = int(sec)
    except (TypeError, ValueError):
        return '-'
    parts = []
    for unit, name in ((86400, '日'), (3600, '時間'), (60, '分')):
        if sec >= unit:
            parts.append(f'{sec // unit}{name}')
            sec %= unit
    if sec or not parts:
        parts.append(f'{sec}秒')
    return ''.join(parts)


# =============================================================================
# xlsx 出力 (標準ライブラリのみで Office Open XML を生成)
# =============================================================================
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
FONT_NAME = '游ゴシック'
_CTRL_RE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _x(s):
    """XML 用エスケープ(XML で使用できない制御文字は除去)"""
    return escape(_CTRL_RE.sub('', str(s)), {'"': '&quot;'})


def col_name(c):
    s = ''
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def cref(r, c):
    return f'{col_name(c)}{r}'


def sheet_ref(sheet, r1, c1, r2=None, c2=None):
    a = f'${col_name(c1)}${r1}'
    if r2 is not None:
        a += f':${col_name(c2)}${r2}'
    return "'" + sheet.replace("'", "''") + "'!" + a


def _num(v):
    if isinstance(v, bool):
        return '1' if v else '0'
    if isinstance(v, int):
        return str(v)
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return repr(float(v))


def _rels(items):
    return (XML_DECL + f'<Relationships xmlns="{NS_PKG_R}">'
            + ''.join(f'<Relationship Id="{i}" Type="{t}" Target="{_x(g)}"/>' for i, t, g in items)
            + '</Relationships>')


class Styles:
    BORDER_COLOR = 'BFBFBF'

    def __init__(self, font_size):
        self.font_size = font_size
        self.fonts, self.fills, self.borders, self.xfs = [], [], [], []
        self.numfmts = {}
        self._index = {}
        self.fills += ['<fill><patternFill patternType="none"/></fill>',
                       '<fill><patternFill patternType="gray125"/></fill>']
        self.get()  # 既定スタイル(index 0)

    def _intern(self, table, key, xml):
        k = (table, key)
        if k not in self._index:
            lst = getattr(self, table)
            self._index[k] = len(lst)
            lst.append(xml)
        return self._index[k]

    def get(self, bold=False, size=None, color=None, fill=None, border=False,
            numfmt=None, halign=None, valign=None, wrap=False):
        key = (bold, size, color, fill, border, numfmt, halign, valign, wrap)
        if ('xfs', key) in self._index:
            return self._index[('xfs', key)]
        font = ('<font>' + ('<b/>' if bold else '') + f'<sz val="{size or self.font_size}"/>'
                + f'<color rgb="FF{color or "000000"}"/>'
                + f'<name val="{FONT_NAME}"/><family val="3"/><charset val="128"/></font>')
        font_id = self._intern('fonts', (bold, size, color), font)
        fill_id = 0
        if fill:
            fill_id = self._intern('fills', fill, '<fill><patternFill patternType="solid">'
                                   f'<fgColor rgb="FF{fill}"/><bgColor indexed="64"/></patternFill></fill>')
        if border:
            side = f'style="thin"><color rgb="FF{self.BORDER_COLOR}"/>'
            border_xml = (f'<border><left {side}</left><right {side}</right>'
                          f'<top {side}</top><bottom {side}</bottom><diagonal/></border>')
        else:
            border_xml = '<border><left/><right/><top/><bottom/><diagonal/></border>'
        border_id = self._intern('borders', border, border_xml)
        numfmt_id = 0
        if numfmt:
            if numfmt not in self.numfmts:
                self.numfmts[numfmt] = 164 + len(self.numfmts)
            numfmt_id = self.numfmts[numfmt]
        align = ''
        if halign or valign or wrap:
            align = ('<alignment' + (f' horizontal="{halign}"' if halign else '')
                     + (f' vertical="{valign}"' if valign else '')
                     + (' wrapText="1"' if wrap else '') + '/>')
        xf = (f'<xf numFmtId="{numfmt_id}" fontId="{font_id}" fillId="{fill_id}" borderId="{border_id}" xfId="0"'
              + (' applyNumberFormat="1"' if numfmt_id else '') + ' applyFont="1"'
              + (' applyFill="1"' if fill_id else '') + (' applyBorder="1"' if border else '')
              + (f' applyAlignment="1">{align}</xf>' if align else '/>'))
        return self._intern('xfs', key, xf)

    def to_xml(self):
        out = [XML_DECL, f'<styleSheet xmlns="{NS_MAIN}">']
        if self.numfmts:
            out.append(f'<numFmts count="{len(self.numfmts)}">'
                       + ''.join(f'<numFmt numFmtId="{i}" formatCode="{_x(c)}"/>' for c, i in self.numfmts.items())
                       + '</numFmts>')
        out.append(f'<fonts count="{len(self.fonts)}">' + ''.join(self.fonts) + '</fonts>')
        out.append(f'<fills count="{len(self.fills)}">' + ''.join(self.fills) + '</fills>')
        out.append(f'<borders count="{len(self.borders)}">' + ''.join(self.borders) + '</borders>')
        out.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        out.append(f'<cellXfs count="{len(self.xfs)}">' + ''.join(self.xfs) + '</cellXfs>')
        out.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        out.append('<dxfs count="0"/><tableStyles count="0"/></styleSheet>')
        return ''.join(out)


class Chart:
    """折れ線(line) / 縦棒(bar) グラフ。系列名は固定文字列、値はセル範囲を参照する。"""

    def __init__(self, kind, title, y_title='件数', cat_fmt='m/d h:mm'):
        self.kind, self.title, self.y_title, self.cat_fmt = kind, title, y_title, cat_fmt
        self.series = []

    def add(self, name, cat_ref, val_ref, color, width=19050):
        self.series.append((name, cat_ref, val_ref, color, width))

    @staticmethod
    def _rich(text, size, bold, rot=None):
        body = f'<a:bodyPr rot="{rot}" vert="horz"/>' if rot is not None else '<a:bodyPr/>'
        b = '1' if bold else '0'
        return (f'<c:title><c:tx><c:rich>{body}<a:lstStyle/><a:p><a:pPr><a:defRPr sz="{size}" b="{b}"/></a:pPr>'
                f'<a:r><a:rPr lang="ja-JP" altLang="en-US" sz="{size}" b="{b}"/><a:t>{_x(text)}</a:t></a:r>'
                '</a:p></c:rich></c:tx><c:overlay val="0"/></c:title>')

    def to_xml(self):
        sers = []
        for i, (name, cat, val, color, width) in enumerate(self.series):
            if self.kind == 'bar':
                fmt = (f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>'
                       '<c:invertIfNegative val="0"/>')
                tail = ''
            else:
                fmt = (f'<c:spPr><a:ln w="{width}" cap="rnd"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                       '<a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>')
                tail = '<c:smooth val="0"/>'
            sers.append(f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/><c:tx><c:v>{_x(name)}</c:v></c:tx>{fmt}'
                        f'<c:cat><c:numRef><c:f>{_x(cat)}</c:f></c:numRef></c:cat>'
                        f'<c:val><c:numRef><c:f>{_x(val)}</c:f></c:numRef></c:val>{tail}</c:ser>')
        if self.kind == 'bar':
            plot = ('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
                    + ''.join(sers)
                    + '<c:gapWidth val="40"/><c:axId val="500001"/><c:axId val="500002"/></c:barChart>')
        else:
            plot = ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>' + ''.join(sers)
                    + '<c:marker val="1"/><c:axId val="500001"/><c:axId val="500002"/></c:lineChart>')
        grid = '<c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
        # 時刻を「テキスト軸」として扱う(auto=0)。日付軸にすると1日単位に丸められるため。
        cat_ax = ('<c:catAx><c:axId val="500001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
                  '<c:delete val="0"/><c:axPos val="b"/>'
                  f'<c:numFmt formatCode="{_x(self.cat_fmt)}" sourceLinked="0"/>'
                  '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>'
                  '<c:txPr><a:bodyPr rot="-2700000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/>'
                  '</a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>'
                  '<c:crossAx val="500002"/><c:crosses val="autoZero"/><c:auto val="0"/><c:lblAlgn val="ctr"/>'
                  '<c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')
        val_ax = ('<c:valAx><c:axId val="500002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling>'
                  f'<c:delete val="0"/><c:axPos val="l"/><c:majorGridlines>{grid}</c:majorGridlines>'
                  + self._rich(self.y_title, 900, False, rot=-5400000)
                  + '<c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="out"/>'
                  '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500001"/>'
                  '<c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>')
        return (XML_DECL + f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
                '<c:roundedCorners val="0"/><c:chart>' + self._rich(self.title, 1200, True)
                + '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' + plot + cat_ax + val_ax + '</c:plotArea>'
                '<c:legend><c:legendPos val="t"/><c:overlay val="0"/></c:legend>'
                '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
                '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
                '<a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
                f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="{FONT_NAME}"/>'
                f'<a:ea typeface="{FONT_NAME}"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>'
                '</c:chartSpace>')


class Sheet:
    def __init__(self, name):
        self.name = name
        self.rows = {}           # row -> {col: (value, style, formula)}
        self.widths = {}
        self.heights = {}
        self.merges = []
        self.freeze = None       # (行, 列): スクロールしない領域の直後のセル
        self.autofilter = None   # (r1, c1, r2, c2)
        self.print_titles = None # (r1, r2)
        self.charts = []
        self.show_grid = True
        self.tab_color = None

    def set(self, r, c, value=None, style=0, formula=None):
        self.rows.setdefault(r, {})[c] = (value, style, formula)

    def merge(self, r1, c1, r2, c2, value=None, style=0):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                self.set(r, c, value if (r, c) == (r1, c1) else None, style)
        self.merges.append(f'{cref(r1, c1)}:{cref(r2, c2)}')

    def add_chart(self, chart, c1, r1, c2, r2):
        """c1,r1-c2,r2 は 0 始まりのセル位置"""
        self.charts.append((chart, (c1, r1, c2, r2)))


class Workbook:
    def __init__(self, font_size=10):
        self.sheets = []
        self.styles = Styles(font_size)
        self.title = ''
        self.creator = ''
        self._sst = {}
        self._sst_list = []

    def add_sheet(self, name):
        sh = Sheet(name)
        self.sheets.append(sh)
        return sh

    def _si(self, s):
        if s not in self._sst:
            self._sst[s] = len(self._sst_list)
            self._sst_list.append(s)
        return self._sst[s]

    def _cell(self, r, c, value, style, formula):
        a = cref(r, c)
        s = f' s="{style}"' if style else ''
        if formula is not None:
            f = f'<f>{_x(formula)}</f>'
            if value is None:
                return f'<c r="{a}"{s}>{f}</c>'
            if isinstance(value, str):
                return f'<c r="{a}"{s} t="str">{f}<v>{_x(value)}</v></c>'
            return f'<c r="{a}"{s}>{f}<v>{_num(value)}</v></c>'
        if value is None or (isinstance(value, float) and not math.isfinite(value)):
            return f'<c r="{a}"{s}/>'
        if isinstance(value, bool):
            return f'<c r="{a}"{s} t="b"><v>{_num(value)}</v></c>'
        if isinstance(value, (int, float)):
            return f'<c r="{a}"{s}><v>{_num(value)}</v></c>'
        return f'<c r="{a}"{s} t="s"><v>{self._si(str(value))}</v></c>'

    def _sheet_xml(self, sh, index, drawing_rid):
        out = [XML_DECL, f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">', '<sheetPr>']
        if sh.tab_color:
            out.append(f'<tabColor rgb="FF{sh.tab_color}"/>')
        out.append('<pageSetUpPr fitToPage="1"/></sheetPr>')
        max_r = max(sh.rows) if sh.rows else 1
        max_c = max((max(cols) for cols in sh.rows.values() if cols), default=1)
        out.append(f'<dimension ref="A1:{cref(max_r, max_c)}"/>')
        view = ('<sheetView workbookViewId="0"' + (' tabSelected="1"' if index == 0 else '')
                + ('' if sh.show_grid else ' showGridLines="0"') + '>')
        if sh.freeze:
            fr, fc = sh.freeze
            top_left = cref(fr, fc)
            if fr > 1 and fc > 1:
                view += (f'<pane xSplit="{fc - 1}" ySplit="{fr - 1}" topLeftCell="{top_left}" activePane="bottomRight" '
                         f'state="frozen"/><selection pane="topRight"/><selection pane="bottomLeft"/>'
                         f'<selection pane="bottomRight" activeCell="{top_left}" sqref="{top_left}"/>')
            elif fr > 1:
                view += (f'<pane ySplit="{fr - 1}" topLeftCell="{top_left}" activePane="bottomLeft" state="frozen"/>'
                         f'<selection pane="bottomLeft" activeCell="{top_left}" sqref="{top_left}"/>')
            elif fc > 1:
                view += (f'<pane xSplit="{fc - 1}" topLeftCell="{top_left}" activePane="topRight" state="frozen"/>'
                         f'<selection pane="topRight" activeCell="{top_left}" sqref="{top_left}"/>')
        out.append(f'<sheetViews>{view}</sheetView></sheetViews>')
        out.append('<sheetFormatPr defaultRowHeight="16.5"/>')
        if sh.widths:
            out.append('<cols>' + ''.join(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>'
                                          for c, w in sorted(sh.widths.items())) + '</cols>')
        out.append('<sheetData>')
        for r in sorted(sh.rows):
            ht = f' ht="{sh.heights[r]}" customHeight="1"' if r in sh.heights else ''
            cells = sh.rows[r]
            out.append(f'<row r="{r}"{ht}>'
                       + ''.join(self._cell(r, c, *cells[c]) for c in sorted(cells)) + '</row>')
        out.append('</sheetData>')
        if sh.autofilter:
            r1, c1, r2, c2 = sh.autofilter
            out.append(f'<autoFilter ref="{cref(r1, c1)}:{cref(r2, c2)}"/>')
        if sh.merges:
            out.append(f'<mergeCells count="{len(sh.merges)}">'
                       + ''.join(f'<mergeCell ref="{m}"/>' for m in sh.merges) + '</mergeCells>')
        out.append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append('<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>')
        if drawing_rid:
            out.append(f'<drawing r:id="{drawing_rid}"/>')
        out.append('</worksheet>')
        return ''.join(out)

    @staticmethod
    def _anchor(i, c1, r1, c2, r2):
        return (f'<xdr:twoCellAnchor editAs="oneCell">'
                f'<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{i + 1}" name="Chart {i}"/>'
                '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                f'<a:graphic><a:graphicData uri="{NS_C}"><c:chart xmlns:c="{NS_C}" xmlns:r="{NS_R}" r:id="rId{i}"/>'
                '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')

    def save(self, path):
        rel_ws = NS_R + '/worksheet'
        rel_drawing = NS_R + '/drawing'
        rel_chart = NS_R + '/chart'
        parts = {}          # パス -> XML
        overrides = []      # (パス, ContentType)
        charts = []
        drawings = 0
        for i, sh in enumerate(self.sheets, 1):
            drawing_rid = None
            if sh.charts:
                drawings += 1
                anchors, rels = [], []
                for j, (chart, (c1, r1, c2, r2)) in enumerate(sh.charts, 1):
                    charts.append(chart)
                    rels.append((f'rId{j}', rel_chart, f'../charts/chart{len(charts)}.xml'))
                    anchors.append(self._anchor(j, c1, r1, c2, r2))
                parts[f'xl/drawings/drawing{drawings}.xml'] = (
                    XML_DECL + f'<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">' + ''.join(anchors) + '</xdr:wsDr>')
                parts[f'xl/drawings/_rels/drawing{drawings}.xml.rels'] = _rels(rels)
                parts[f'xl/worksheets/_rels/sheet{i}.xml.rels'] = _rels(
                    [('rId1', rel_drawing, f'../drawings/drawing{drawings}.xml')])
                overrides.append((f'/xl/drawings/drawing{drawings}.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'))
                drawing_rid = 'rId1'
            parts[f'xl/worksheets/sheet{i}.xml'] = self._sheet_xml(sh, i - 1, drawing_rid)
            overrides.append((f'/xl/worksheets/sheet{i}.xml',
                              'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'))
        for n, chart in enumerate(charts, 1):
            parts[f'xl/charts/chart{n}.xml'] = chart.to_xml()
            overrides.append((f'/xl/charts/chart{n}.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'))

        # 共有文字列(シート XML 生成後に確定)
        parts['xl/sharedStrings.xml'] = (
            XML_DECL + f'<sst xmlns="{NS_MAIN}" count="{len(self._sst_list)}" uniqueCount="{len(self._sst_list)}">'
            + ''.join(f'<si><t xml:space="preserve">{_x(s)}</t></si>' for s in self._sst_list) + '</sst>')
        parts['xl/styles.xml'] = self.styles.to_xml()

        defined = []
        for i, sh in enumerate(self.sheets):
            if sh.autofilter:
                r1, c1, r2, c2 = sh.autofilter
                defined.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
                               f'{_x(sheet_ref(sh.name, r1, c1, r2, c2))}</definedName>')
            if sh.print_titles:
                q = "'" + sh.name.replace("'", "''") + "'"
                defined.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{i}">'
                               f'{_x(q)}!${sh.print_titles[0]}:${sh.print_titles[1]}</definedName>')
        parts['xl/workbook.xml'] = (
            XML_DECL + f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><workbookPr/>'
            '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15600" activeTab="0"/></bookViews>'
            '<sheets>' + ''.join(f'<sheet name="{_x(sh.name)}" sheetId="{i}" r:id="rId{i}"/>'
                                 for i, sh in enumerate(self.sheets, 1)) + '</sheets>'
            + (f'<definedNames>{"".join(defined)}</definedNames>' if defined else '')
            + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        n = len(self.sheets)
        parts['xl/_rels/workbook.xml.rels'] = _rels(
            [(f'rId{i}', rel_ws, f'worksheets/sheet{i}.xml') for i in range(1, n + 1)]
            + [(f'rId{n + 1}', NS_R + '/styles', 'styles.xml'),
               (f'rId{n + 2}', NS_R + '/sharedStrings', 'sharedStrings.xml')])
        now = dt.datetime.now(dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        parts['docProps/core.xml'] = (
            XML_DECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{_x(self.title)}</dc:title><dc:creator>{_x(self.creator)}</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
        parts['docProps/app.xml'] = (
            XML_DECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>sqs_metrics_report</Application></Properties>')
        parts['_rels/.rels'] = _rels([
            ('rId1', NS_R + '/officeDocument', 'xl/workbook.xml'),
            ('rId2', NS_PKG_R + '/metadata/core-properties', 'docProps/core.xml'),
            ('rId3', NS_R + '/extended-properties', 'docProps/app.xml')])
        overrides += [
            ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
            ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
            ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
            ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
            ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')]
        content_types = (
            XML_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + ''.join(f'<Override PartName="{p}" ContentType="{t}"/>' for p, t in overrides) + '</Types>')

        tmp = path + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', content_types.encode('utf-8'))
            for name in ['_rels/.rels', 'docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml',
                         'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml']:
                z.writestr(name, parts.pop(name).encode('utf-8'))
            for name in sorted(parts):
                z.writestr(name, parts[name].encode('utf-8'))
        os.replace(tmp, path)


# =============================================================================
# メトリクスの読み込み・集計
# =============================================================================
def load_results(path):
    """get-metric-data の結果を {Id: {epoch: 値}} に変換(ページ分割された結果も統合)"""
    series, notes = {}, []
    if not path or not os.path.exists(path):
        return series, notes
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    for res in data.get('MetricDataResults', []):
        d = series.setdefault(res['Id'], {})
        for t, v in zip(res.get('Timestamps', []), res.get('Values', [])):
            d[parse_ts(t)] = float(v)
        for m in res.get('Messages') or []:
            notes.append(f'CloudWatch({res["Id"]}): {m.get("Code")} {m.get("Value")}')
    for m in data.get('Messages') or []:
        notes.append(f'CloudWatch: {m.get("Code")} {m.get("Value")}')
    return series, notes


def regrid(results, prefix, origin, period):
    """メトリクスを詳細区間(origin から period 秒刻み)に揃える"""
    q = {}
    for key, _m, _s, how in METRICS:
        g = {}
        for ts, v in results.get(f'{prefix}_{key}', {}).items():
            slot = origin + ((ts - origin) // period) * period
            if slot in g:
                g[slot] = g[slot] + v if how == 'sum' else max(g[slot], v)
            else:
                g[slot] = v
        q[key] = g
    return q


def aggregate_hourly(q, start, n_hours):
    rows = [{'slots': set()} for _ in range(n_hours)]
    for key, _m, _s, how in METRICS:
        for ts in sorted(q[key]):
            h = (ts - start) // HOUR
            if ts < start or h >= n_hours:
                continue
            v = q[key][ts]
            r = rows[h]
            r['slots'].add(ts)
            if how == 'sum':
                r[key] = r.get(key, 0.0) + v
            else:
                r[key] = max(r.get(key, v), v)
            if key == 'visible':
                r['visible_end'] = v
    for r in rows:
        r['points'] = len(r.pop('slots'))
        for key, _m, _s, how in METRICS:
            if how == 'sum':
                r.setdefault(key, 0.0)
        r['size_avg'] = r['size_sum'] / r['size_cnt'] if r['size_cnt'] else None
    return rows


def stock_flow(q, base_results, prefix, start, end, period):
    """総メッセージ数(可視+処理中+遅延)の推移から、詳細区間ごとの増加分(=追加数の推定値)を算出"""
    series = {}
    for k in STOCK_KEYS:
        for ts, v in q[k].items():
            series[ts] = series.get(ts, 0.0) + v
    # 期間開始時点の基準値: 直前1時間の詳細データ → 開始前14日間の1時間データ → 0 とみなす
    base_val, base_ts, base_src = 0.0, None, 'assumed'
    bkeys = [f'b{prefix}_{k}' for k in STOCK_KEYS]
    bts = sorted({t for k in bkeys for t in base_results.get(k, {}) if t < start})
    if bts:
        base_ts = bts[-1]
        base_val = sum(base_results.get(k, {}).get(base_ts, 0.0) for k in bkeys)
        base_src = 'hourly'
    pre = sorted(t for t in series if t < start)
    if pre:
        base_ts, base_val, base_src = pre[-1], series[pre[-1]], 'detail'
    events = {}
    prev_val, prev_ts = base_val, base_ts
    for ts in sorted(t for t in series if start <= t < end):
        total = series[ts]
        events[ts] = {
            'total': total, 'prev': prev_val, 'prev_ts': prev_ts,
            'delta': total - prev_val, 'inflow': max(0.0, total - prev_val),
            # 直前の値が無い / 欠損をはさんだ値(欠損期間中の追加をまとめて計上)
            'gap': prev_ts is None or ts - prev_ts > 2 * period,
        }
        prev_val, prev_ts = total, ts
    return SimpleNamespace(events=events, base_val=base_val, base_src=base_src,
                           end_val=prev_val, has_data=bool(events))


def hourly_flow(flow, start, n_hours):
    rows = [{'inflow': 0.0, 'count': 0, 'gap': False, 'end': None} for _ in range(n_hours)]
    for ts in sorted(flow.events):
        e = flow.events[ts]
        r = rows[(ts - start) // HOUR]
        if e['inflow'] > 0:
            r['inflow'] += e['inflow']
            r['count'] += 1
            r['gap'] = r['gap'] or e['gap']
        r['end'] = e['total']
    carry = flow.base_val
    for r in rows:
        r['begin'] = carry
        r['change'] = r['end'] - carry if r['end'] is not None else None
        if r['end'] is not None:
            carry = r['end']
    return rows


def peak_index(values):
    best = None
    for i, v in enumerate(values):
        if v is not None and v > 0 and (best is None or v > values[best]):
            best = i
    return best


# =============================================================================
# レポート作成
# =============================================================================
C_MAIN, C_DLQ = '2F5597', 'C00000'


def make_styles(wb):
    S = wb.styles.get
    hdr = dict(bold=True, color='FFFFFF', border=True, halign='center', valign='center', wrap=True)
    tot = dict(bold=True, fill='F2F2F2', border=True)
    return {
        'title': S(bold=True, size=16, color='1F3864'),
        'subtitle': S(size=10, color='404040'),
        'note': S(size=9, color='595959'),
        'section': S(bold=True, size=11, color='FFFFFF', fill='1F4E78', valign='center'),
        'th': S(fill='44546A', **hdr),
        'th_main': S(fill=C_MAIN, **hdr),
        'th_dlq': S(fill=C_DLQ, **hdr),
        'key': S(bold=True, fill='DDEBF7', border=True, valign='center', wrap=True),
        'text': S(border=True, valign='center'),
        'wrap': S(border=True, valign='center', wrap=True),
        'center': S(border=True, halign='center', valign='center'),
        'int': S(border=True, numfmt='#,##0', valign='center'),
        'dec': S(border=True, numfmt='#,##0.0', valign='center'),
        'pct': S(border=True, numfmt='0.00%', valign='center'),
        'dt': S(border=True, numfmt='yyyy/mm/dd hh:mm', halign='center', valign='center'),
        'int_alert': S(bold=True, color='9C0006', fill='FFC7CE', border=True, numfmt='#,##0', valign='center'),
        'text_alert': S(color='9C0006', fill='FFC7CE', border=True, valign='center'),
        'dt_alert': S(color='9C0006', fill='FFC7CE', border=True, numfmt='yyyy/mm/dd hh:mm', halign='center',
                      valign='center'),
        'muted': S(color='808080', border=True, valign='center'),
        'tot_label': S(halign='center', **tot),
        'tot_int': S(numfmt='#,##0', **tot),
        'tot_dec': S(numfmt='#,##0.0', **tot),
        'tot_pct': S(numfmt='0.00%', **tot),
        'tot_blank': S(fill='F2F2F2', border=True),
        'judge_info': S(bold=True, size=12, color='1F3864', fill='DDEBF7', border=True, valign='center', wrap=True),
        'judge_ng': S(bold=True, size=12, color='9C0006', fill='FFC7CE', border=True, valign='center', wrap=True),
        'judge_ok': S(bold=True, size=12, color='006100', fill='C6EFCE', border=True, valign='center', wrap=True),
        'judge_na': S(bold=True, size=12, color='7F6000', fill='FFEB9C', border=True, valign='center', wrap=True),
    }


def section(ws, st, r, title, c2=4):
    ws.merge(r, 1, r, c2, title, st['section'])
    ws.heights[r] = 21
    return r + 1


def write_table(ws, st, hr, columns, rows, totals=None, freeze_col=1, header_style='th'):
    """columns: [(見出し, 列幅, スタイル名)]、rows の要素は値 or (値, スタイル名)
    totals: 列ごとに None / 'sum' / 'max' / ('label', 文字) / ('value', 値, スタイル名) / ('formula', 式, 値, スタイル名)
    戻り値: (データ開始行, データ終了行, 合計行)"""
    for j, (head, width, _style) in enumerate(columns, 1):
        hs = header_style[j - 1] if isinstance(header_style, (list, tuple)) else header_style
        ws.set(hr, j, head, st[hs])
        ws.widths[j] = width
    ws.heights[hr] = 36
    r = hr
    for row in rows:
        r += 1
        for j, v in enumerate(row, 1):
            style = columns[j - 1][2]
            if isinstance(v, tuple):
                v, style = v
            ws.set(r, j, v, st[style])
    first, last = hr + 1, r
    ws.freeze = (hr + 1, freeze_col + 1)
    ws.print_titles = (hr, hr)
    if rows:
        ws.autofilter = (hr, 1, last, len(columns))
    total_row = None
    if totals and rows:
        total_row = last + 1
        for j, spec in enumerate(totals, 1):
            rng = f'{cref(first, j)}:{cref(last, j)}'
            nums = [v[0] if isinstance(v, tuple) else v for v in (row[j - 1] for row in rows)]
            nums = [v for v in nums if isinstance(v, (int, float))]
            if spec == 'sum':
                ws.set(total_row, j, sum(nums), st['tot_int'], f'SUBTOTAL(109,{rng})')
            elif spec == 'max':
                ws.set(total_row, j, max(nums) if nums else 0, st['tot_int'], f'SUBTOTAL(104,{rng})')
            elif isinstance(spec, tuple) and spec[0] == 'label':
                ws.set(total_row, j, spec[1], st['tot_label'])
            elif isinstance(spec, tuple) and spec[0] == 'value':
                ws.set(total_row, j, spec[1], st[spec[2]])
            elif isinstance(spec, tuple) and spec[0] == 'formula':
                ws.set(total_row, j, spec[2], st[spec[3]], spec[1])
            else:
                ws.set(total_row, j, None, st['tot_blank'])
    return first, last, total_row


def build_summary(wb, st, c):
    ws = wb.add_sheet('サマリー')
    ws.show_grid = False
    ws.tab_color = '1F4E78'
    ws.widths = {1: 38, 2: 40, 3: 40, 4: 56}
    has_dlq = c.dlq is not None
    ws.set(1, 1, 'SQS メッセージ追加状況レポート', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'メインキュー: {c.main_name}　／　DLQ: {c.dlq_name or "なし"}　／　'
                 f'集計期間: {c.period_label}（{c.tz_label}）', st['subtitle'])

    # ---- 判定
    r = section(ws, st, 4, '■ 判定')
    mt = c.main_total
    mp = c.main_peak
    main_line = f'メインキュー追加数 合計: {mt:,.0f} 件'
    if mp is not None:
        main_line += f'（最大 {c.main_hourly[mp]["sent"]:,.0f} 件／時: {hour_label(c.hours[mp])}）'
    ws.merge(r, 1, r, 4, main_line, st['judge_info'])
    ws.heights[r] = 24
    r += 1
    if not has_dlq:
        dlq_line, style = 'DLQ: 未設定のため集計していません', st['judge_na']
    elif c.dlq_total > 0:
        dp = c.dlq_peak
        dlq_line = (f'DLQ への追加あり: 合計 {c.dlq_total:,.0f} 件（推定）／ 発生時間帯 {c.dlq_hours_with_inflow} 件'
                    f' ／ 最大 {c.dlq_hourly[dp]["inflow"]:,.0f} 件／時: {hour_label(c.hours[dp])}')
        style = st['judge_ng']
    elif not c.dlq_flow.has_data:
        dlq_line, style = 'DLQ への追加なし（期間中 DLQ のメトリクスなし＝非アクティブ）', st['judge_ok']
    else:
        dlq_line, style = 'DLQ への追加なし（推定 0 件）', st['judge_ok']
    ws.merge(r, 1, r, 4, dlq_line, style)
    ws.heights[r] = 24
    r += 2

    # ---- 実行条件
    r = section(ws, st, r, '■ 実行条件')
    dlq_src = {'auto': 'メインキューの RedrivePolicy から自動検出', 'arg': '引数(-d)で指定',
               'none': '集計対象外(-d none)', 'notfound': 'RedrivePolicy 未設定のため集計なし'}.get(c.dlq_source, '-')
    conds = [
        ('対象キュー', c.main_name),
        ('DLQ', f'{c.dlq_name}（{dlq_src}）' if c.dlq_name else dlq_src),
        ('リージョン', c.region or '-'),
        ('AWS アカウント ID', c.account or '-'),
        ('集計期間', c.period_label),
        ('タイムゾーン', c.tz_label),
        ('集計単位', '1時間（各行は開始時刻から1時間）'),
        ('詳細データ粒度', f'{c.period}秒（{fmt_duration(c.period)}）'),
        ('レポート作成日時', fmt_time(c.now, '%Y/%m/%d %H:%M:%S')),
        ('作成ホスト', c.host),
        ('作成ツール', f'sqs_metrics_report.sh v{c.version}'),
    ]
    for k, v in conds:
        ws.set(r, 1, k, st['key'])
        ws.merge(r, 2, r, 4, v, st['text'])
        r += 1
    r += 1

    # ---- 期間集計
    r = section(ws, st, r, '■ 期間集計')
    for j, (h, s) in enumerate([('項目', 'th'), ('メインキュー', 'th_main'), ('DLQ', 'th_dlq'), ('備考', 'th')], 1):
        ws.set(r, j, h, st[s])
    r += 1
    mh, dh = c.main_hourly, c.dlq_hourly
    n = c.n_hours

    def gmax(rows, key):
        vals = [x.get(key) for x in rows]
        i = peak_index(vals)
        return (vals[i], i) if i is not None else (0 if any(v is not None for v in vals) else None, None)

    m_vis, m_vis_i = gmax(mh, 'visible')
    m_age, m_age_i = gmax(mh, 'age')
    stats = [
        ('追加メッセージ数 合計', mt, c.dlq_total if has_dlq else None, 'int',
         'メイン: NumberOfMessagesSent ／ DLQ: 総メッセージ数の増加分から推定'),
        ('1時間あたり平均追加数', mt / n, c.dlq_total / n if has_dlq else None, 'dec', f'全 {n} 時間で平均'),
        ('1時間あたり最大追加数', mh[mp]['sent'] if mp is not None else 0,
         (dh[c.dlq_peak]['inflow'] if c.dlq_peak is not None else 0) if has_dlq else None, 'int', ''),
        ('最大追加時間帯', hour_label(c.hours[mp]) if mp is not None else '-',
         (hour_label(c.hours[c.dlq_peak]) if c.dlq_peak is not None else '-') if has_dlq else None, 'center', ''),
        ('追加が発生した時間帯数', sum(1 for x in mh if x['sent'] > 0),
         c.dlq_hours_with_inflow if has_dlq else None, 'int', f'全 {n} 時間中'),
        ('受信数 合計', sum(x['received'] for x in mh),
         sum(x['received'] for x in c.dlq_hourly_m) if has_dlq else None, 'int', 'NumberOfMessagesReceived（再受信を含む）'),
        ('削除数 合計', sum(x['deleted'] for x in mh),
         sum(x['deleted'] for x in c.dlq_hourly_m) if has_dlq else None, 'int', 'NumberOfMessagesDeleted'),
        ('NumberOfMessagesSent 合計', mt,
         sum(x['sent'] for x in c.dlq_hourly_m) if has_dlq else None, 'int',
         'DLQ: 手動送信のみ（maxReceiveCount 超過による自動移動は計上されない）'),
        ('滞留数（可視）最大', m_vis, gmax(c.dlq_hourly_m, 'visible')[0] if has_dlq else None, 'int',
         f'メイン最大時間帯: {hour_label(c.hours[m_vis_i])}' if m_vis_i is not None else ''),
        ('最古メッセージ経過秒 最大', m_age, gmax(c.dlq_hourly_m, 'age')[0] if has_dlq else None, 'int',
         f'メイン: {fmt_duration(m_age)}' if m_age else ''),
        ('期間開始時点の総メッセージ数', c.main_flow.base_val if c.main_flow.base_src != 'assumed' else '-',
         c.dlq_flow.base_val if has_dlq else None, 'int',
         '可視+処理中+遅延' + ('（DLQ: 開始前14日間にデータが無いため 0 とみなして算出）'
                          if has_dlq and c.dlq_flow.base_src == 'assumed' else '')),
        ('期間終了時点の総メッセージ数', c.main_flow.end_val if c.main_flow.has_data else '-',
         c.dlq_flow.end_val if has_dlq else None, 'int', '期間内の最終データ時点'),
        ('データ欠損時間帯数', sum(1 for x in mh if x['points'] == 0),
         sum(1 for x in c.dlq_hourly_m if x['points'] == 0) if has_dlq else None, 'int',
         'メトリクスが無い時間帯（6時間以上非アクティブなキュー等）'),
    ]
    for label, mv, dv, style, note in stats:
        ws.set(r, 1, label, st['key'])
        ws.set(r, 2, mv, st[style] if isinstance(mv, (int, float)) else st['center'])
        if has_dlq:
            alert = label == '追加メッセージ数 合計' and dv and dv > 0
            ws.set(r, 3, dv, st['int_alert'] if alert else (st[style] if isinstance(dv, (int, float)) else st['center']))
        else:
            ws.set(r, 3, '-', st['center'])
        ws.set(r, 4, note, st['wrap'])
        r += 1
    ws.set(r, 1, 'DLQ 移動率（DLQ 追加数 ÷ メイン追加数）', st['key'])
    rate = c.dlq_total / mt if has_dlq and mt else None
    ws.merge(r, 2, r, 3, rate if rate is not None else '-', st['pct'] if rate is not None else st['center'])
    ws.set(r, 4, '同一期間内での比率（参考値）', st['wrap'])
    r += 2

    # ---- キュー設定
    r = section(ws, st, r, '■ キュー設定（レポート作成時点）')
    for j, (h, s) in enumerate([('項目', 'th'), ('メインキュー', 'th_main'), ('DLQ', 'th_dlq'), ('備考', 'th')], 1):
        ws.set(r, j, h, st[s])
    r += 1
    ma, da = c.main_attrs, c.dlq_attrs

    def attr_rows(at, name, url):
        rp = parse_redrive(at)
        created = at.get('CreatedTimestamp')
        sse = '-'
        if at.get('KmsMasterKeyId'):
            sse = f'SSE-KMS ({at["KmsMasterKeyId"]})'
        elif at.get('SqsManagedSseEnabled') == 'true':
            sse = 'SSE-SQS'
        elif at:
            sse = '無効'

        def num(k):
            v = at.get(k)
            return int(v) if v is not None and str(v).isdigit() else '-'
        return [
            name or '-', url or '-', at.get('QueueArn', '-'),
            ('FIFO' if at.get('FifoQueue') == 'true' else '標準') if at else '-',
            fmt_time(int(created), '%Y/%m/%d %H:%M:%S') if created and str(created).isdigit() else '-',
            fmt_duration(at['VisibilityTimeout']) if 'VisibilityTimeout' in at else '-',
            fmt_duration(at['MessageRetentionPeriod']) if 'MessageRetentionPeriod' in at else '-',
            fmt_duration(at['DelaySeconds']) if 'DelaySeconds' in at else '-',
            f'{int(at["MaximumMessageSize"]):,} byte' if str(at.get('MaximumMessageSize', '')).isdigit() else '-',
            fmt_duration(at['ReceiveMessageWaitTimeSeconds']) if 'ReceiveMessageWaitTimeSeconds' in at else '-',
            rp.get('deadLetterTargetArn', '-') if at else '-',
            str(rp.get('maxReceiveCount', '-')) if at else '-',
            sse,
            num('ApproximateNumberOfMessages'),
            num('ApproximateNumberOfMessagesNotVisible'),
            num('ApproximateNumberOfMessagesDelayed'),
        ]
    labels = ['キュー名', 'キュー URL', 'キュー ARN', 'キュー種別', '作成日時', '可視性タイムアウト',
              'メッセージ保持期間', '配信遅延', '最大メッセージサイズ', '受信メッセージ待機時間',
              'DLQ（deadLetterTargetArn）', '最大受信数（maxReceiveCount）', '暗号化',
              '現在のメッセージ数（可視）', '現在のメッセージ数（処理中）', '現在のメッセージ数（遅延）']
    notes = {'DLQ（deadLetterTargetArn）': 'メインキューの RedrivePolicy',
             '最大受信数（maxReceiveCount）': 'この回数を超えて受信されたメッセージが DLQ へ移動',
             '現在のメッセージ数（可視）': 'レポート作成時点の ApproximateNumberOfMessages'}
    mrow = attr_rows(ma, c.main_name, c.main_url)
    drow = attr_rows(da, c.dlq_name, c.dlq_url) if has_dlq else ['-'] * len(labels)
    for label, mv, dv in zip(labels, mrow, drow):
        ws.set(r, 1, label, st['key'])
        ws.set(r, 2, mv, st['int'] if isinstance(mv, int) else st['wrap'])
        ws.set(r, 3, dv, st['int'] if isinstance(dv, int) else st['wrap'])
        ws.set(r, 4, notes.get(label, ''), st['wrap'])
        r += 1
    r += 1

    # ---- 注意事項・警告
    r = section(ws, st, r, '■ 注意事項・警告')
    for note in (c.notes or ['特になし']):
        ws.merge(r, 1, r, 4, note, st['wrap'])
        r += 1
    r += 1

    # ---- シート構成
    r = section(ws, st, r, '■ シート構成')
    for name, desc in c.sheet_list:
        ws.set(r, 1, name, st['key'])
        ws.merge(r, 2, r, 4, desc, st['wrap'])
        r += 1


def build_overview(wb, st, c):
    ws = wb.add_sheet('時間別レポート')
    ws.tab_color = C_MAIN
    has_dlq = c.dlq is not None
    ws.set(1, 1, '時間別 メッセージ追加状況', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'メインキュー: {c.main_name}　／　DLQ: {c.dlq_name or "なし"}　／　'
                 f'集計期間: {c.period_label}（{c.tz_label}）', st['subtitle'])
    ws.set(3, 1, '※ 各行は「開始日時」から1時間の集計値です。DLQ 追加数は総メッセージ数の増加分から算出した推定値です'
                 '（算出方法は「指標の説明」シートを参照）。合計行はフィルター後の表示行のみを集計します。', st['note'])
    gr, hr = 5, 6
    cols = [('開始日時', 17, 'dt'), ('終了日時', 17, 'dt'),
            ('追加数', 11, 'int'), ('受信数', 11, 'int'), ('削除数', 11, 'int'),
            ('滞留数\n(可視・最大)', 15, 'int'), ('最古メッセージ\n経過秒(最大)', 17, 'int')]
    hstyles = ['th', 'th'] + ['th_main'] * 5
    if has_dlq:
        cols += [('追加数\n(推定)', 11, 'int'), ('総メッセージ数\n(時間末)', 17, 'int'), ('DLQ移動率', 13, 'pct')]
        hstyles += ['th_dlq'] * 3
    cols += [('備考', 44, 'text')]
    hstyles += ['th']
    nc = len(cols)
    ws.merge(gr, 1, gr, 2, '時間帯', st['th'])
    ws.merge(gr, 3, gr, 7, f'メインキュー: {c.main_name}', st['th_main'])
    if has_dlq:
        ws.merge(gr, 8, gr, 10, f'DLQ: {c.dlq_name}', st['th_dlq'])
    ws.set(gr, nc, None, st['th'])
    ws.heights[gr] = 20

    rows = []
    for i, hs in enumerate(c.hours):
        m = c.main_hourly[i]
        notes = []
        if m['points'] == 0:
            notes.append('メインキュー: データなし')
        row = [xl_time(hs), xl_time(hs + HOUR), m['sent'], m['received'], m['deleted'],
               m.get('visible'), m.get('age')]
        if has_dlq:
            d = c.dlq_hourly[i]
            rate = d['inflow'] / m['sent'] if m['sent'] else None
            row += [(d['inflow'], 'int_alert') if d['inflow'] > 0 else d['inflow'], d['end'], rate]
            if d['inflow'] > 0:
                notes.insert(0, 'DLQ 追加あり' + ('（データ欠損後の初回値を含む）' if d['gap'] else ''))
        note = ' ／ '.join(notes)
        row.append((note, 'text_alert') if has_dlq and c.dlq_hourly[i]['inflow'] > 0 else note)
        rows.append(row)

    totals = [('label', '合計 / 最大'), ('value', None, 'tot_label'), 'sum', 'sum', 'sum', 'max', 'max']
    if has_dlq:
        tr = hr + len(rows) + 1
        rate = c.dlq_total / c.main_total if c.main_total else '-'
        totals += ['sum', ('value', c.dlq_flow.end_val, 'tot_int'),
                   ('formula', f'IF(C{tr}>0,H{tr}/C{tr},"-")', rate, 'tot_pct')]
    totals += [None]
    first, last, total_row = write_table(ws, st, hr, cols, rows, totals, freeze_col=2, header_style=hstyles)
    ws.merges.append(f'A{total_row}:B{total_row}')
    ws.print_titles = (gr, hr)
    return SimpleNamespace(name=ws.name, first=first, last=last)


def build_charts(wb, st, c, ov):
    ws = wb.add_sheet('グラフ')
    ws.show_grid = False
    ws.tab_color = C_MAIN
    ws.set(1, 1, '時間別グラフ', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　データ: 「時間別レポート」シート', st['subtitle'])
    cat = sheet_ref(ov.name, ov.first, 1, ov.last, 1)

    def col(n):
        return sheet_ref(ov.name, ov.first, n, ov.last, n)

    ch = Chart('line', f'メインキュー 時間別メッセージ数（{c.main_name}）')
    ch.add('追加数', cat, col(3), C_MAIN, width=34925)
    ch.add('受信数', cat, col(4), 'ED7D31')
    ch.add('削除数', cat, col(5), '70AD47')
    ws.add_chart(ch, 0, 3, 20, 24)
    top = 26
    if c.dlq is not None:
        ch = Chart('bar', f'DLQ 時間別追加数（推定）（{c.dlq_name}）')
        ch.add('DLQ 追加数（推定）', cat, col(8), C_DLQ)
        ws.add_chart(ch, 0, top, 20, top + 21)
        top += 23
    ch = Chart('line', '滞留状況')
    ch.add('メインキュー 滞留数（可視・最大）', cat, col(6), C_MAIN)
    if c.dlq is not None:
        ch.add('DLQ 総メッセージ数（時間末）', cat, col(9), C_DLQ)
    ws.add_chart(ch, 0, top, 20, top + 21)


def build_hourly_main(wb, st, c):
    ws = wb.add_sheet('時間別_メインキュー')
    ws.tab_color = C_MAIN
    ws.set(1, 1, f'時間別 メインキュー詳細（{c.main_name}）', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　詳細データ粒度: {c.period}秒　'
                 f'※ 件数はデータが無い時間帯を 0、滞留数等は空欄で表示', st['subtitle'])
    hr = 4
    cols = [('開始日時', 17, 'dt'), ('終了日時', 17, 'dt'), ('追加数', 11, 'int'), ('受信数', 11, 'int'),
            ('削除数', 11, 'int'), ('空受信数', 12, 'int'), ('滞留数\n(可視・最大)', 15, 'int'),
            ('滞留数\n(可視・時間末)', 16, 'int'), ('処理中\n(最大)', 11, 'int'), ('遅延中\n(最大)', 11, 'int'),
            ('最古メッセージ\n経過秒(最大)', 17, 'int'), ('平均メッセージ\nサイズ(byte)', 17, 'dec'),
            ('最大メッセージ\nサイズ(byte)', 17, 'int'), ('データ点数', 13, 'int')]
    rows = []
    for i, hs in enumerate(c.hours):
        m = c.main_hourly[i]
        rows.append([xl_time(hs), xl_time(hs + HOUR), m['sent'], m['received'], m['deleted'], m['empty'],
                     m.get('visible'), m.get('visible_end'), m.get('notvisible'), m.get('delayed'), m.get('age'),
                     m['size_avg'], m.get('size_max'), m['points']])
    size_cnt = sum(m['size_cnt'] for m in c.main_hourly)
    avg = sum(m['size_sum'] for m in c.main_hourly) / size_cnt if size_cnt else None
    totals = [('label', '合計 / 最大'), ('value', None, 'tot_label'), 'sum', 'sum', 'sum', 'sum', 'max', None,
              'max', 'max', 'max', ('value', avg, 'tot_dec'), 'max', 'sum']
    _f, _l, tr = write_table(ws, st, hr, cols, rows, totals, freeze_col=2, header_style='th_main')
    ws.merges.append(f'A{tr}:B{tr}')


def build_hourly_dlq(wb, st, c):
    ws = wb.add_sheet('時間別_DLQ')
    ws.tab_color = C_DLQ
    ws.set(1, 1, f'時間別 DLQ 詳細（{c.dlq_name}）', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　詳細データ粒度: {c.period}秒　'
                 '※ 追加数(推定) = 詳細区間ごとの総メッセージ数(可視+処理中+遅延)の増加分の合計', st['subtitle'])
    hr = 4
    cols = [('開始日時', 17, 'dt'), ('終了日時', 17, 'dt'), ('追加数\n(推定)', 11, 'int'), ('追加検知\n回数', 11, 'int'),
            ('総メッセージ数\n(時間初)', 18, 'int'), ('総メッセージ数\n(時間末)', 18, 'int'), ('増減', 10, 'int'),
            ('滞留数\n(可視・最大)', 15, 'int'), ('処理中\n(最大)', 11, 'int'), ('遅延中\n(最大)', 11, 'int'),
            ('手動送信数\n(Sent)', 13, 'int'), ('受信数', 11, 'int'), ('削除数', 11, 'int'),
            ('最古メッセージ\n経過秒(最大)', 17, 'int'), ('データ点数', 13, 'int'), ('備考', 60, 'text')]
    rows = []
    for i, hs in enumerate(c.hours):
        d, m = c.dlq_hourly[i], c.dlq_hourly_m[i]
        alert = d['inflow'] > 0
        note = ''
        if alert and d['gap']:
            note = 'データ欠損後の初回値を含む（欠損期間中の追加をまとめて計上）'
        elif m['points'] == 0:
            note = 'データなし（非アクティブ）'
        rows.append([xl_time(hs), xl_time(hs + HOUR), (d['inflow'], 'int_alert') if alert else d['inflow'], d['count'],
                     d['begin'], d['end'], d['change'], m.get('visible'), m.get('notvisible'), m.get('delayed'),
                     m['sent'], m['received'], m['deleted'], m.get('age'), m['points'],
                     (note, 'text_alert') if alert else note])
    tr = hr + len(rows) + 1
    totals = [('label', '合計 / 最大'), ('value', None, 'tot_label'), 'sum', 'sum',
              ('value', c.dlq_flow.base_val, 'tot_int'), ('value', c.dlq_flow.end_val, 'tot_int'),
              ('formula', f'F{tr}-E{tr}', c.dlq_flow.end_val - c.dlq_flow.base_val, 'tot_int'),
              'max', 'max', 'max', 'sum', 'sum', 'sum', 'max', 'sum', None]
    _f, _l, tr = write_table(ws, st, hr, cols, rows, totals, freeze_col=2, header_style='th_dlq')
    ws.merges.append(f'A{tr}:B{tr}')


def build_dlq_events(wb, st, c):
    ws = wb.add_sheet('DLQ追加一覧')
    ws.tab_color = C_DLQ
    ws.set(1, 1, f'DLQ 追加検知一覧（{c.dlq_name}）', st['title'])
    ws.heights[1] = 28
    events = [(ts, e) for ts, e in sorted(c.dlq_flow.events.items()) if e['inflow'] > 0]
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　詳細データ粒度: {c.period}秒　'
                 f'検知 {len(events)} 回 ／ 追加数(推定) 合計 {c.dlq_total:,.0f} 件', st['subtitle'])
    ws.set(3, 1, '※ 総メッセージ数(可視+処理中+遅延)が直前のデータより増加した時刻を「追加」として検知しています。'
                 'メインキューの値は同じ詳細区間の値です。', st['note'])
    hr = 5
    cols = [('No', 7, 'int'), ('検知日時', 17, 'dt'), ('追加数\n(推定)', 11, 'int_alert'),
            ('直前の\n総メッセージ数', 19, 'int'), ('直前データ日時', 17, 'dt'), ('追加後の\n総メッセージ数', 19, 'int'),
            ('DLQ 最古メッセージ\n経過秒', 21, 'int'), ('メインキュー\n受信数(同区間)', 19, 'int'),
            ('メインキュー\n滞留数(同区間)', 19, 'int'), ('備考', 60, 'wrap')]
    if not events:
        for j, (head, width, _s) in enumerate(cols, 1):
            ws.set(hr, j, head, st['th_dlq'])
            ws.widths[j] = width
        ws.heights[hr] = 36
        ws.merge(hr + 1, 1, hr + 1, len(cols), '対象期間内に DLQ への追加は検知されませんでした。', st['judge_ok'])
        return
    rows = []
    mq, dq = c.main, c.dlq
    for n, (ts, e) in enumerate(events, 1):
        note = ''
        if e['prev_ts'] is None:
            note = '期間開始前のデータが無いため、直前の総メッセージ数を 0 とみなして算出'
        elif e['gap']:
            note = f'{fmt_time(e["prev_ts"])} 以降データ欠損（欠損期間中の追加をまとめて計上）'
        rows.append([n, xl_time(ts), e['inflow'], e['prev'],
                     xl_time(e['prev_ts']) if e['prev_ts'] is not None else None, e['total'],
                     dq['age'].get(ts), mq['received'].get(ts), mq['visible'].get(ts), note])
    totals = [('label', '合計'), ('value', None, 'tot_label'), 'sum'] + [None] * 7
    _f, _l, tr = write_table(ws, st, hr, cols, rows, totals, freeze_col=2, header_style='th_dlq')
    ws.merges.append(f'A{tr}:B{tr}')


def build_detail_main(wb, st, c):
    ws = wb.add_sheet('詳細_メインキュー')
    ws.tab_color = '8EA9DB'
    ws.set(1, 1, f'詳細データ メインキュー（{c.main_name}）', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　粒度: {c.period}秒（各行は日時から {fmt_duration(c.period)} の値）'
                 '　※ 空欄は CloudWatch にデータが無いことを示します', st['subtitle'])
    hr = 4
    cols = [('日時', 17, 'dt'), ('追加数', 10, 'int'), ('受信数', 10, 'int'), ('削除数', 10, 'int'),
            ('空受信数', 12, 'int'), ('滞留数\n(可視)', 11, 'int'), ('処理中', 10, 'int'), ('遅延中', 10, 'int'),
            ('最古メッセージ\n経過秒', 17, 'int'), ('平均メッセージ\nサイズ(byte)', 17, 'dec')]
    q = c.main
    rows = []
    for ts in range(c.start, c.end, c.period):
        cnt = q['size_cnt'].get(ts)
        rows.append([xl_time(ts), q['sent'].get(ts), q['received'].get(ts), q['deleted'].get(ts),
                     q['empty'].get(ts), q['visible'].get(ts), q['notvisible'].get(ts), q['delayed'].get(ts),
                     q['age'].get(ts), q['size_sum'].get(ts, 0.0) / cnt if cnt else None])
    write_table(ws, st, hr, cols, rows, freeze_col=1, header_style='th_main')


def build_detail_dlq(wb, st, c):
    ws = wb.add_sheet('詳細_DLQ')
    ws.tab_color = 'F4B183'
    ws.set(1, 1, f'詳細データ DLQ（{c.dlq_name}）', st['title'])
    ws.heights[1] = 28
    ws.set(2, 1, f'集計期間: {c.period_label}（{c.tz_label}）　粒度: {c.period}秒　'
                 '※ 追加数(推定) = 総メッセージ数の直前データからの増加分　空欄は CloudWatch にデータが無いことを示します',
           st['subtitle'])
    hr = 4
    cols = [('日時', 17, 'dt'), ('追加数\n(推定)', 11, 'int'), ('総メッセージ数', 17, 'int'), ('増減', 9, 'int'),
            ('滞留数\n(可視)', 11, 'int'), ('処理中', 10, 'int'), ('遅延中', 10, 'int'),
            ('最古メッセージ\n経過秒', 17, 'int'), ('手動送信数\n(Sent)', 13, 'int'), ('受信数', 10, 'int'),
            ('削除数', 10, 'int'), ('備考', 50, 'text')]
    q, events = c.dlq, c.dlq_flow.events
    rows = []
    for ts in range(c.start, c.end, c.period):
        e = events.get(ts)
        inflow = total = delta = None
        note = ''
        if e:
            inflow, total, delta = e['inflow'], e['total'], e['delta']
            if inflow > 0 and e['gap']:
                note = ('期間開始前データなし（直前値を 0 とみなす）' if e['prev_ts'] is None
                        else f'{fmt_time(e["prev_ts"])} 以降データ欠損後の値')
        alert = bool(inflow)
        rows.append([(xl_time(ts), 'dt_alert') if alert else xl_time(ts),
                     (inflow, 'int_alert') if alert else inflow, total, delta,
                     q['visible'].get(ts), q['notvisible'].get(ts), q['delayed'].get(ts), q['age'].get(ts),
                     q['sent'].get(ts), q['received'].get(ts), q['deleted'].get(ts),
                     (note, 'text_alert') if alert else note])
    write_table(ws, st, hr, cols, rows, freeze_col=1, header_style='th_dlq')


def build_glossary(wb, st, c):
    ws = wb.add_sheet('指標の説明')
    ws.show_grid = False
    ws.widths = {1: 30, 2: 40, 3: 16, 4: 90}
    ws.set(1, 1, '指標の説明・注意事項', st['title'])
    ws.heights[1] = 28
    r = section(ws, st, 3, '■ 指標')
    for j, h in enumerate(['レポート上の項目', 'CloudWatch メトリクス', '統計', '説明'], 1):
        ws.set(r, j, h, st['th'])
    r += 1
    items = [
        ('追加数', 'NumberOfMessagesSent', 'Sum',
         'キューへ正常に追加(SendMessage / SendMessageBatch)されたメッセージ数。SNS・EventBridge 等からの配信も含みます。'),
        ('受信数', 'NumberOfMessagesReceived', 'Sum', 'ReceiveMessage で返されたメッセージ数。同一メッセージの再受信も含みます。'),
        ('削除数', 'NumberOfMessagesDeleted', 'Sum', '削除に成功した件数。同一メッセージの重複削除も含みます。'),
        ('空受信数', 'NumberOfEmptyReceives', 'Sum', 'メッセージを返さなかった ReceiveMessage の呼び出し数。'),
        ('滞留数(可視)', 'ApproximateNumberOfMessagesVisible', 'Maximum', '取得可能な状態のメッセージ数(概算)。'),
        ('処理中', 'ApproximateNumberOfMessagesNotVisible', 'Maximum', '受信済みで未削除(可視性タイムアウト中)のメッセージ数(概算)。'),
        ('遅延中', 'ApproximateNumberOfMessagesDelayed', 'Maximum', '配信遅延中で取得できないメッセージ数(概算)。'),
        ('最古メッセージ経過秒', 'ApproximateAgeOfOldestMessage', 'Maximum',
         '最も古い未処理メッセージの経過秒数。DLQ では DLQ へ移動された時刻からの経過秒数になります。'),
        ('平均 / 最大メッセージサイズ', 'SentMessageSize', 'Sum÷SampleCount / Maximum', '追加されたメッセージのサイズ(byte)。'),
        ('総メッセージ数', '(算出) 可視 + 処理中 + 遅延', '-', 'キュー内に存在するメッセージの総数(概算)。'),
        ('DLQ 追加数(推定)', '(算出)', '-',
         '詳細区間ごとに DLQ の総メッセージ数が直前のデータから増加した分を合計した値。'
         'maxReceiveCount 超過による自動移動と、DLQ への手動送信の両方を含みます。'),
        ('DLQ 移動率', '(算出)', '-',
         'DLQ 追加数(推定) ÷ メインキュー追加数。同じ時間帯に追加されたメッセージが DLQ へ移動したとは限らないため参考値です。'),
        ('時間末 / 時間初', '(算出)', '-', '各時間帯の最後のデータ / 直前の時間帯までの最後のデータの値。'),
        ('データ点数', '-', '-', f'その時間帯で CloudWatch にデータが存在した詳細区間の数(最大 {HOUR // c.period})。'),
    ]
    for item in items:
        for j, v in enumerate(item, 1):
            ws.set(r, j, v, st['key'] if j == 1 else st['wrap'])
        r += 1
    r += 1
    r = section(ws, st, r, '■ 注意事項')
    notes = [
        'DLQ への自動移動(maxReceiveCount 超過)は DLQ の NumberOfMessagesSent に計上されません(AWS の仕様)。'
        'そのため本レポートでは DLQ の総メッセージ数(可視+処理中+遅延)の増加分から DLQ への追加数を推定しています。',
        '同じ詳細区間内で DLQ への追加と DLQ からの削除(再処理・Redrive 等)が同時に発生した場合、推定値は実際より小さくなります。'
        '詳細データの粒度が細かいほど精度が上がります(1分粒度は直近15日以内のみ取得可能)。',
        'SQS のメトリクスは1分間隔で CloudWatch に送信されます。キューが6時間以上非アクティブ(メッセージなし・API 操作なし)の場合、'
        'メトリクスは送信されずデータが欠損します。欠損した時間帯の件数は 0、滞留数等は空欄で表示しています。',
        '欠損後に最初に得られた値と欠損前の値との差分は、欠損期間中の追加としてまとめて計上します(「DLQ追加一覧」の備考に表示)。'
        'また、非アクティブ状態から復帰した直後のメトリクスは最大15分程度遅れて記録されることがあります。',
        '期間開始時点の総メッセージ数は、開始直前1時間の詳細データ、無い場合は開始前14日間の1時間粒度データの最終値を使用します。'
        'いずれも無い場合は 0 とみなします(SQS のメッセージ保持期間は最大14日)。',
        'CloudWatch のデータ保持期間: 1分粒度 = 15日、5分粒度 = 63日、1時間粒度 = 455日。'
        '粒度は -g オプション(auto の場合は開始日時から自動選択)で指定します。',
        'Approximate 系のメトリクスは SQS の分散アーキテクチャによる概算値です。',
        '本スクリプトの GetQueueAttributes 呼び出しにより、キューは CloudWatch 上でアクティブ扱いになります(過去のメトリクス値には影響しません)。',
    ]
    for i, n in enumerate(notes, 1):
        ws.set(r, 1, f'注意 {i}', st['key'])
        ws.merge(r, 2, r, 4, n, st['wrap'])
        ws.heights[r] = 36
        r += 1


def cmd_report(a):
    set_tz(a.tz, a.tz_offset)
    start, end, period = a.start, a.end, a.period
    detail, notes1 = load_results(a.metrics)
    baseline, notes2 = load_results(a.baseline)
    n_hours = (end - start) // HOUR
    hours = [start + i * HOUR for i in range(n_hours)]
    off = local(start).strftime('%z')

    c = SimpleNamespace()
    c.start, c.end, c.period, c.now = start, end, period, a.now
    c.n_hours, c.hours = n_hours, hours
    c.tz_label = f'{a.tz} (UTC{off[:3]}:{off[3:]})'
    c.period_label = f'{fmt_time(start)} ～ {fmt_time(end)}（{n_hours}時間）'
    c.main_name, c.main_url = a.main_name, a.main_url
    c.dlq_name, c.dlq_url, c.dlq_source = a.dlq_name, a.dlq_url, a.dlq_source
    c.main_attrs = load_attrs(a.main_attrs)
    c.dlq_attrs = load_attrs(a.dlq_attrs)
    arn = c.main_attrs.get('QueueArn', '').split(':')
    c.account = arn[4] if len(arn) > 5 else ''
    c.region = a.region or (arn[3] if len(arn) > 5 else '')
    c.host, c.version = a.host, a.version
    c.notes = list(a.note) + notes1 + notes2

    # 詳細区間は期間開始時刻を起点に揃える(開始前の先読みデータも同じ刻みになる)
    c.main = regrid(detail, 'm', start, period)
    c.main_hourly = aggregate_hourly(c.main, start, n_hours)
    c.main_flow = stock_flow(c.main, baseline, 'm', start, end, period)
    c.main_total = sum(m['sent'] for m in c.main_hourly)
    c.main_peak = peak_index([m['sent'] for m in c.main_hourly])
    if not any(c.main[k] for k in c.main):
        c.notes.append('メインキューのメトリクスが取得できませんでした(キュー名・リージョン・期間・データ保持期間を確認してください)')

    c.dlq = c.dlq_hourly = c.dlq_hourly_m = c.dlq_flow = c.dlq_peak = None
    c.dlq_total, c.dlq_hours_with_inflow = 0.0, 0
    if a.dlq_name:
        c.dlq = regrid(detail, 'd', start, period)
        c.dlq_hourly_m = aggregate_hourly(c.dlq, start, n_hours)
        c.dlq_flow = stock_flow(c.dlq, baseline, 'd', start, end, period)
        c.dlq_hourly = hourly_flow(c.dlq_flow, start, n_hours)
        c.dlq_total = sum(d['inflow'] for d in c.dlq_hourly)
        c.dlq_peak = peak_index([d['inflow'] for d in c.dlq_hourly])
        c.dlq_hours_with_inflow = sum(1 for d in c.dlq_hourly if d['inflow'] > 0)
        if c.dlq_flow.base_src == 'assumed' and c.dlq_flow.has_data:
            c.notes.append('DLQ: 期間開始前14日間にメトリクスが無いため、期間開始時点の総メッセージ数を 0 とみなして追加数を算出しました')

    c.sheet_list = [
        ('サマリー', '判定・実行条件・期間集計・キュー設定・注意事項'),
        ('時間別レポート', '1時間ごとのメインキュー追加数と DLQ 追加数(推定)の一覧'),
        ('グラフ', '時間別の追加数・DLQ 追加数・滞留状況のグラフ'),
        ('時間別_メインキュー', 'メインキューの全メトリクスの1時間集計'),
    ]
    if c.dlq is not None:
        c.sheet_list += [('時間別_DLQ', 'DLQ の全メトリクスと追加数(推定)の1時間集計'),
                         ('DLQ追加一覧', 'DLQ への追加を検知した時刻の一覧(詳細粒度)')]
    c.sheet_list.append(('詳細_メインキュー', f'メインキューの {c.period}秒粒度の全データ'))
    if c.dlq is not None:
        c.sheet_list.append(('詳細_DLQ', f'DLQ の {c.period}秒粒度の全データと追加数(推定)'))
    c.sheet_list.append(('指標の説明', '各指標の定義・算出方法・注意事項'))

    wb = Workbook()
    wb.title = f'SQS メッセージ追加状況レポート ({c.main_name})'
    wb.creator = 'sqs_metrics_report.sh'
    st = make_styles(wb)
    build_summary(wb, st, c)
    ov = build_overview(wb, st, c)
    build_charts(wb, st, c, ov)
    build_hourly_main(wb, st, c)
    if c.dlq is not None:
        build_hourly_dlq(wb, st, c)
        build_dlq_events(wb, st, c)
    build_detail_main(wb, st, c)
    if c.dlq is not None:
        build_detail_dlq(wb, st, c)
    build_glossary(wb, st, c)
    wb.save(a.output)

    print(f'  メインキュー追加数 合計 : {c.main_total:,.0f} 件')
    if c.dlq is not None:
        print(f'  DLQ 追加数(推定) 合計  : {c.dlq_total:,.0f} 件（発生時間帯 {c.dlq_hours_with_inflow} 件）')


def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest='cmd')
    q = sub.add_parser('queries')
    q.add_argument('--kind', choices=['detail', 'baseline'], required=True)
    q.add_argument('--period', type=int, required=True)
    q.add_argument('--main', required=True)
    q.add_argument('--dlq')
    q = sub.add_parser('attr')
    q.add_argument('--attrs', required=True)
    q.add_argument('--name', required=True)
    q = sub.add_parser('dlq-name')
    q.add_argument('--attrs', required=True)
    q = sub.add_parser('report')
    for name in ('output', 'main-name', 'main-url', 'main-attrs', 'metrics', 'tz', 'version'):
        q.add_argument(f'--{name}', required=True)
    for name in ('start', 'end', 'now', 'period'):
        q.add_argument(f'--{name}', type=int, required=True)
    for name in ('baseline', 'tz-offset', 'region', 'dlq-source', 'host', 'dlq-name', 'dlq-url', 'dlq-attrs'):
        q.add_argument(f'--{name}', default='')
    q.add_argument('--note', action='append', default=[])
    a = p.parse_args()
    {'queries': cmd_queries, 'attr': cmd_attr, 'dlq-name': cmd_dlq_name, 'report': cmd_report}[a.cmd](a)


if __name__ == '__main__':
    main()
PYEOF
}

#------------------------------------------------------------------------------
# メイン処理
#------------------------------------------------------------------------------
main() {
  parse_args "$@"
  check_prereqs
  resolve_times
  resolve_output

  WORK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/sqs_report.XXXXXX")" || die 2 "作業ディレクトリを作成できません"
  trap cleanup EXIT
  trap 'exit 130' INT
  trap 'exit 143' TERM
  write_python

  resolve_queues
  fetch_metrics
  build_report
}

main "$@"
