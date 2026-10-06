#!/bin/bash
# ==============================================================================
# KỊCH BẢN 1-CLICK TRIỂN KHAI LÊN GOOGLE CLOUD RUN
# ==============================================================================
SERVICE_NAME="chem-docx-engine"
REGION="asia-southeast1" # Singapore (gần Việt Nam nhất, tốc độ cao nhất)

echo "🚀 [1/3] Bắt đầu triển khai dịch vụ $SERVICE_NAME lên Google Cloud Run ($REGION)..."
gcloud run deploy $SERVICE_NAME \
  --source . \
  --platform managed \
  --region $REGION \
  --allow-unauthenticated \
  --memory 512Mi \
  --cpu 1 \
  --min-instances 0 \
  --max-instances 2 \
  --port 8080

echo "✅ [2/3] Triển khai thành công!"
echo "👉 Hãy sao chép Service URL ở trên (dạng https://...run.app) và dán vào nút Cấu Hình của Web App!"
