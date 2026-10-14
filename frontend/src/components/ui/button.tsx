import { Slot } from '@radix-ui/react-slot'
import { cva, type VariantProps } from 'class-variance-authority'
import type { ButtonHTMLAttributes } from 'react'
import { cn } from '@/lib/utils'

const buttonVariants = cva(
  'inline-flex h-9 cursor-pointer items-center justify-center gap-2 whitespace-nowrap rounded-lg border px-4 text-sm font-medium transition-[background-color,transform] duration-150 ease-out active:scale-[0.98] disabled:pointer-events-none disabled:opacity-50 motion-reduce:transition-none motion-reduce:active:scale-100',
  {
    variants: {
      variant: {
        primary: 'border-accent bg-accent text-on-accent hover:bg-accent-strong hover:border-accent-strong',
        secondary: 'border-border bg-surface text-text hover:bg-surface-2',
        ink: 'border-ink bg-ink text-ink-fg hover:opacity-90',
        ghost: 'border-transparent bg-transparent text-muted hover:bg-surface-2 hover:text-text',
      },
      size: { default: '', sm: 'h-8 px-3' },
    },
    defaultVariants: { variant: 'secondary', size: 'default' },
  },
)

export interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement>, VariantProps<typeof buttonVariants> {
  asChild?: boolean
}

export function Button({ className, variant, size, asChild, ...props }: ButtonProps) {
  const Comp = asChild ? Slot : 'button'
  return <Comp className={cn(buttonVariants({ variant, size }), className)} {...props} />
}
