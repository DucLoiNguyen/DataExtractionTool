# Dùng dự án này để học DevOps

Ứng dụng web (`webapp/`) là một dịch vụ nhỏ, có trạng thái tạm (tệp tải lên), có endpoint sức khoẻ, đủ để tập toàn bộ vòng
đời: kiểm thử → đóng gói → CI → chạy container → mô tả triển khai. **Chưa có gì được triển khai**; mọi thứ chạy cục bộ.

## Có gì trong repo

| Thành phần | Tệp | Học được gì |
|---|---|---|
| Kiểm thử tự động | `tests/`, `pyproject.toml` | pytest, `TestClient`, kiểm thử tích hợp không mock |
| Lint | `ruff` (`pyproject.toml`) | Quy tắc chất lượng chạy trong CI |
| Đóng gói | `Dockerfile`, `.dockerignore` | Multi-stage build, user không phải root, `HEALTHCHECK`, `ARG` tuỳ chọn |
| Chạy cục bộ | `docker-compose.yml` | Root filesystem chỉ đọc, `tmpfs`, bỏ capability |
| CI | `.github/workflows/ci.yml` | Job `test` → job `docker` (build + smoke test), cache, quyền tối thiểu |
| Mô tả triển khai | `deploy/k8s/deployment.yaml` | Deployment + Service, probe, resource limit, securityContext |
| Tiện ích | `Makefile` | Gom lệnh hay dùng |

## Lộ trình gợi ý

1. **Chạy test và lint**: `make install && make lint && make test`. Phá thử một test rồi xem CI đỏ ra sao.
2. **Docker**: `make docker-build && make docker-run`, mở http://localhost:8000. Thử `docker history`, `docker image ls`
   để so kích thước; đổi `--build-arg WITH_LIBREOFFICE=1` và so sánh.
3. **Compose**: `make compose-up`. Thử bỏ `tmpfs` rồi xem lỗi vì sao ứng dụng cần ghi `/tmp`.
4. **CI**: đẩy nhánh lên GitHub, xem tab Actions. Sau đó thử thêm bước: quét lỗ hổng ảnh (Trivy), đẩy ảnh lên GHCR.
5. **Kubernetes** (kind/minikube): `kind load docker-image dataextraction-web:local` rồi `kubectl apply -f deploy/k8s/`.
   Thử `kubectl rollout`, tắt một pod, sửa probe sai để thấy pod không Ready.

## Bài tập tiếp theo (chưa làm sẵn, để bạn tự làm)

- Đẩy ảnh lên registry với tag theo commit SHA, thêm job deploy chạy khi merge vào `main`.
- Thêm Ingress + TLS, hoặc deploy lên một dịch vụ container được quản lý (Cloud Run, ECS, Fly.io...).
- Viết hạ tầng bằng Terraform; thêm `/metrics` (Prometheus) và dashboard Grafana.
- Thêm xác thực (ứng dụng hiện **không có đăng nhập**; tệp tải lên chứa dữ liệu cá nhân nên đừng mở ra Internet).

## Lưu ý đã biết

- Mỗi yêu cầu trích xuất chạy tuần tự (khoá toàn cục), nên nhiều bản sao (`replicas`) tăng độ sẵn sàng chứ chưa tăng thông lượng nhiều.
- Kết quả giữ trong bộ nhớ/thư mục tạm của từng pod 1 giờ: nếu tải xuống rơi vào pod khác sẽ báo "hết hạn".
  Đây là bài học tốt về ứng dụng có trạng thái (cần sticky session hoặc lưu kết quả ra volume/object storage).
- Dockerfile và manifest k8s mới được kiểm tra cú pháp YAML, chưa build/chạy thực tế trong môi trường tạo ra chúng.
