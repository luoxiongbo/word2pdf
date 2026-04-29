# Cloud Run Deployment (Low Cost)

This project can run as a public web service with Cloud Run.

## Why Cloud Run

- Pay-as-you-go
- Scale to zero when idle
- Public URL by default (`https://<service>-<hash>-<region>.run.app`)
- Good fit for this repo because LibreOffice runs inside the same container

## Prerequisites

- Google Cloud project
- Billing enabled
- `gcloud` CLI installed
- Logged in:

```bash
gcloud auth login
gcloud auth application-default login
```

## First-time setup

```bash
gcloud services enable \
  run.googleapis.com \
  cloudbuild.googleapis.com \
  artifactregistry.googleapis.com
```

## Deploy

```bash
PROJECT_ID="your-gcp-project-id" ./scripts/deploy_cloud_run.sh
```

Defaults in `scripts/deploy_cloud_run.sh`:
- Region: `us-central1`
- CPU / memory: `1 vCPU / 1Gi`
- Concurrency: `1`
- Timeout: `300s`
- Min instances: `0`
- Max instances: `2`
- Upload limit env: `MAX_UPLOAD_MB=30`

## Cost guardrails

- Keep `--min-instances 0` to avoid idle charges
- Keep `--max-instances` low for budget control
- Use `concurrency=1` for stable document conversion
- Keep upload size below Cloud Run request limit (32 MiB); this repo default is 30 MiB in cloud

## Update deployment

After code changes:

```bash
PROJECT_ID="your-gcp-project-id" ./scripts/deploy_cloud_run.sh
```

## Optional: custom domain

Cloud Run works directly with `run.app` URL.
If you need your own domain, map a custom domain in Cloud Run console.
