type CardItem = {
  id: string;
  row: Record<string, string | number | boolean | null>;
  meta: {
    title: string;
    badge: { tone: 'default' | 'low' | 'out' | 'info'; label: string } | null;
  };
};

type InventoryCardGridProps = {
  headers: string[];
  rows: CardItem[];
};

export function InventoryCardGrid({ headers, rows }: InventoryCardGridProps) {
  if (!rows.length) {
    return null;
  }

  return (
    <div className="grid">
      {rows.map((item) => {
        const badgeClass = item.meta.badge?.tone === 'low'
          ? 'badge low-stock'
          : item.meta.badge?.tone === 'out'
            ? 'badge out-of-stock'
            : 'badge';

        return (
          <article key={item.id} className="card">
            <div className="meta">{headers.length ? `${headers.length} columns` : 'Parsed row'}</div>
            {item.meta.badge ? <span className={badgeClass}>{item.meta.badge.label}</span> : null}
            <h3 className="card-title">{item.meta.title}</h3>
            <div className="field-list">
              {headers.map((header) => {
                const value = item.row[header];
                if (value === null || value === undefined || value === '') return null;
                return (
                  <div key={`${item.id}-${header}`} className="field-row">
                    <strong>{header}</strong>
                    <span>{String(value)}</span>
                  </div>
                );
              })}
            </div>
          </article>
        );
      })}
    </div>
  );
}
