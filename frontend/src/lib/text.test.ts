import { describe, expect, it } from 'vitest'
import { compareText, fold } from './text'

describe('fold', () => {
  it('bo dau va dau d', () => {
    expect(fold('Nguyễn Đức Lợi')).toBe('nguyen duc loi')
  })
  it('chiu duoc gia tri rong', () => {
    expect(fold(null)).toBe('')
    expect(fold(undefined)).toBe('')
  })
  it('gop to hop NFD', () => {
    expect(fold('xã Mai Châu')).toBe('xa mai chau')
  })
})

describe('compareText', () => {
  it('xep so theo gia tri', () => {
    expect(['10', '2', '1'].sort(compareText)).toEqual(['1', '2', '10'])
  })
  it('khong phan biet dau', () => {
    expect(compareText('An', 'Ăn')).toBeLessThanOrEqual(0)
  })
})
