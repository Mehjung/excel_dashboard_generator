import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.gridspec import GridSpec
from PIL import Image as PILImage

# Beispieldaten erstellen
data = {
    'Datum': ['2024-01-01', '2024-01-01', '2024-02-01', '2024-02-01', '2024-03-01', '2024-03-01'],
    'Produkt': ['Produkt A', 'Produkt B', 'Produkt A', 'Produkt B', 'Produkt A', 'Produkt B'],
    'Region': ['Nord', 'Süd', 'Nord', 'Süd', 'Nord', 'Süd'],
    'Umsatz': [100, 200, 150, 250, 200, 300]
}

df = pd.DataFrame(data)

# Setze den Stil auf 'whitegrid'
sns.set(style="whitegrid")

# Erstellen der Figur und des Layouts
fig = plt.figure(figsize=(18, 12))
gs = GridSpec(2, 2, figure=fig, height_ratios=[2, 1])

# Balkendiagramm
ax1 = fig.add_subplot(gs[0, 0])
sns.barplot(data=df.groupby('Datum')['Umsatz'].sum().reset_index(), x='Datum', y='Umsatz', ax=ax1, palette="Blues_d")
ax1.set_title('Umsatz nach Monat', fontsize=16)
ax1.set_xlabel('Datum', fontsize=12)
ax1.set_ylabel('Umsatz', fontsize=12)

# Kreisdiagramm
ax2 = fig.add_subplot(gs[0, 1])
df.groupby('Produkt')['Umsatz'].sum().plot(kind='pie', autopct='%1.1f%%', colors=sns.color_palette("pastel"), ax=ax2)
ax2.set_title('Umsatzverteilung nach Produkt', fontsize=16)
ax2.set_ylabel('')

# Liniendiagramm
ax3 = fig.add_subplot(gs[1, :])
sns.lineplot(data=df, x='Datum', y='Umsatz', hue='Produkt', marker='o', palette="husl", ax=ax3)
ax3.set_title('Umsatzentwicklung nach Produkt', fontsize=16)
ax3.set_xlabel('Datum', fontsize=12)
ax3.set_ylabel('Umsatz', fontsize=12)

# Layout-Anpassungen
plt.tight_layout()
plt.savefig('/mnt/data/improved_dashboard.png')
plt.show()
