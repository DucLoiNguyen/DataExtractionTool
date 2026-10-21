import { animate, m, useMotionValue, useReducedMotion, useTransform } from 'motion/react'
import { useEffect } from 'react'
import { fmt } from '@/lib/utils'

/** So dem len tu 0 (chi la hieu ung nho; tat khi nguoi dung chon giam chuyen dong). */
export function CountUp({ value, className }: { value: number; className?: string }) {
  const reduce = useReducedMotion()
  const mv = useMotionValue(reduce ? value : 0)
  const text = useTransform(mv, (v) => fmt(Math.round(v)))

  useEffect(() => {
    if (reduce) {
      mv.set(value)
      return
    }
    const controls = animate(mv, value, { duration: 0.6, ease: 'easeOut' })
    return () => controls.stop()
  }, [value, reduce, mv])

  return <m.span className={className}>{text}</m.span>
}
