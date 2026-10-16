import * as TabsPrimitive from '@radix-ui/react-tabs'
import type { ComponentProps } from 'react'
import { cn } from '@/lib/utils'

export const Tabs = TabsPrimitive.Root

export function TabsList({ className, ...props }: ComponentProps<typeof TabsPrimitive.List>) {
  return <TabsPrimitive.List className={cn('flex flex-wrap gap-2', className)} {...props} />
}

export function TabsTrigger({ className, ...props }: ComponentProps<typeof TabsPrimitive.Trigger>) {
  return (
    <TabsPrimitive.Trigger
      className={cn(
        'cursor-pointer rounded-lg border border-border bg-surface px-4 py-2 text-sm text-text transition-colors duration-150',
        'hover:bg-surface-2 data-[state=active]:border-accent data-[state=active]:bg-accent data-[state=active]:text-on-accent',
        'motion-reduce:transition-none',
        className,
      )}
      {...props}
    />
  )
}

export const TabsContent = TabsPrimitive.Content
