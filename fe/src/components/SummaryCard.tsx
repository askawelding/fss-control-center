import type { LucideIcon } from 'lucide-react'

export type SummaryCardProps = {
  tag: string
  label: string
  value: string
  note: string
  icon: LucideIcon
  variant: 'orange' | 'purple' | 'gray'
}

export default function SummaryCard({
  tag,
  label,
  value,
  note,
  icon: Icon,
  variant,
}: SummaryCardProps) {
  return (
    <article className="stat-card">
      <div className="stat-card-top">
        <div className={`stat-icon icon-${variant}`}>
          <Icon size={20} />
        </div>

        <span className="stat-tag">{tag}</span>
      </div>

      <p className="stat-label">{label}</p>

      <h3>{value}</h3>

      <p className="stat-note">{note}</p>
    </article>
  )
}