import * as React from 'react';
import { ChevronDown } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Label } from '@/components/ui/label';
import { Switch } from '@/components/ui/switch';
import { Tooltip, TooltipPopup, TooltipTrigger } from '@/components/ui/tooltip';
import { cn } from '@/lib/utils';

/** A cluster of ribbon commands, separated from its neighbours by a hairline. */
export function RibbonGroup({
  children,
  className,
  label,
}: {
  children: React.ReactNode;
  className?: string;
  /** Accessible group name (PowerPoint shows it under the group in the classic ribbon). */
  label?: string;
}) {
  return (
    <div
      aria-label={label}
      className={cn(
        'flex shrink-0 items-center gap-0.5 px-1.5 first:ps-0 [&+&]:border-s',
        className,
      )}
      role="group"
    >
      {children}
    </div>
  );
}

type RibbonButtonProps = Omit<React.ComponentProps<typeof Button>, 'size' | 'variant'> & {
  /** Accessible name and tooltip text. */
  label: string;
  /** Optional keyboard hint shown in the tooltip. */
  shortcut?: string;
  /** Render the label next to the icon. */
  showLabel?: boolean;
  active?: boolean;
};

/** Flat, PowerPoint-style command button. Icon-only by default, with a tooltip. */
export function RibbonButton({
  active,
  children,
  className,
  label,
  shortcut,
  showLabel,
  ...props
}: RibbonButtonProps) {
  return (
    <Tooltip>
      <TooltipTrigger
        render={
          <Button
            aria-label={label}
            aria-pressed={active}
            className={cn(
              "h-7 min-w-7 rounded-md px-1.5 font-normal text-[13px] sm:h-7 [&_svg:not([class*='size-'])]:size-4",
              showLabel && 'px-2',
              active && 'bg-brand/10 text-brand hover:bg-brand/15',
              className,
            )}
            size="sm"
            variant="ghost"
            {...props}
          />
        }
      >
        {children}
        {showLabel ? <span>{label}</span> : null}
      </TooltipTrigger>
      <TooltipPopup side="bottom">
        {label}
        {shortcut ? <span className="ms-2 text-muted-foreground">{shortcut}</span> : null}
      </TooltipPopup>
    </Tooltip>
  );
}

/** Ghost dropdown trigger used for ribbon menus. */
export function RibbonMenuButton({
  className,
  icon,
  label,
  showLabel = true,
  ...props
}: React.ComponentProps<typeof Button> & {
  label: string;
  icon?: React.ReactNode;
  showLabel?: boolean;
}) {
  return (
    <Button
      aria-label={label}
      className={cn(
        "h-7 gap-1 rounded-md px-1.5 font-normal text-[13px] sm:h-7 [&_svg:not([class*='size-'])]:size-4",
        className,
      )}
      size="sm"
      variant="ghost"
      {...props}
    >
      {icon}
      {showLabel ? <span>{label}</span> : null}
      <ChevronDown className="size-3! opacity-60" />
    </Button>
  );
}

/** A labelled switch that fits the single-row ribbon. */
export function RibbonToggle({
  checked,
  disabled,
  label,
  onCheckedChange,
}: {
  checked: boolean;
  disabled?: boolean;
  label: string;
  onCheckedChange: (checked: boolean) => void;
}) {
  return (
    <Label className="h-7 cursor-pointer gap-2 rounded-md px-1.5 text-[13px] font-normal hover:bg-accent has-disabled:cursor-default has-disabled:opacity-64 has-disabled:hover:bg-transparent">
      <Switch
        checked={checked}
        className="data-checked:bg-brand"
        disabled={disabled}
        onCheckedChange={onCheckedChange}
      />
      {label}
    </Label>
  );
}

export function RibbonDivider() {
  return <div aria-hidden className="mx-1 h-5 w-px shrink-0 bg-border" />;
}
