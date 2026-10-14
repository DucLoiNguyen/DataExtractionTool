
import openpyxl
import pytest



@pytest.fixture
def sample_xlsx(tmp_path):
    """Tep Excel nho: 5 dong, 1 thieu email, 1 trung email, 1 email go sai (gmail,com)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DS1"
    ws.append(["STT", "Họ và tên", "Đơn vị công tác", "Email", "Số điện thoại"])
    ws.append([1, "nguyễn văn an", "UBND phường Ninh Kiều", "an.nv@cantho.gov.vn", "912345678"])
    ws.append([2, "Trần Thị Bình", "Sở Tư pháp", "binh@cantho.gov.vn", "0987654321"])
    ws.append([3, "Lê Văn Cường", "Sở Tư pháp", "cuong@gmail,com", "0901234567"])
    ws.append([4, "Phạm Dung", "Sở Tư pháp", "", "0901234568"])
    ws.append([5, "Phạm Dung", "Sở Tư pháp", "binh@cantho.gov.vn", "0901234569"])
    path = tmp_path / "mau.xlsx"
    wb.save(path)
    return path
