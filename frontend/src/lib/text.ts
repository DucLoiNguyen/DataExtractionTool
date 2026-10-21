/** Bo dau tieng Viet + ha chu thuong, de tim kiem khong can go dau. "Đặng" -> "dang". */
export function fold(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
}

/** So sanh theo thu tu chu cai tieng Viet, so duoc xep theo gia tri ("2" < "10"). */
export function compareText(a: unknown, b: unknown): number {
  return fold(a).localeCompare(fold(b), 'vi', { numeric: true })
}
