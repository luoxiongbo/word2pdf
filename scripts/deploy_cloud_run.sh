#!/usr/bin/env bash
set -euo pipefail

# Usage:
#   PROJECT_ID=your-gcp-project-id ./scripts/deploy_cloud_run.sh
#
# Optional env vars:
#   SERVICE_NAME=word-to-pdf
#   REGION=us-central1
#   MAX_UPLOAD_MB=30
#   CONVERT_TIMEOUT_SECONDS=180
#   MAX_INSTANCES=2

PROJECT_ID="${PROJECT_ID:-}"
SERVICE_NAME="${SERVICE_NAME:-word-to-pdf}"
REGION="${REGION:-us-central1}"
MAX_UPLOAD_MB="${MAX_UPLOAD_MB:-30}"
CONVERT_TIMEOUT_SECONDS="${CONVERT_TIMEOUT_SECONDS:-180}"
MAX_INSTANCES="${MAX_INSTANCES:-2}"

if [[ -z "${PROJECT_ID}" ]]; then
  echo "PROJECT_ID is required. Example:"
  echo "  PROJECT_ID=my-gcp-project ./scripts/deploy_cloud_run.sh"
  exit 1
fi

gcloud config set project "${PROJECT_ID}"

gcloud run deploy "${SERVICE_NAME}" \
  --source . \
  --region "${REGION}" \
  --allow-unauthenticated \
  --port 8080 \
  --memory 1Gi \
  --cpu 1 \
  --concurrency 1 \
  --timeout 300 \
  --min-instances 0 \
  --max-instances "${MAX_INSTANCES}" \
  --set-env-vars "MAX_UPLOAD_MB=${MAX_UPLOAD_MB},CONVERT_TIMEOUT_SECONDS=${CONVERT_TIMEOUT_SECONDS}"

echo "Deploy complete."
