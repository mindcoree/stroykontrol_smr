"""Offline geographic context for the inspection hexbin (no API key or tiles).

The local equirectangular projection uses kilometres at 41.85° N. Its small
distortion over Chicago is sufficient for this descriptive overview. Counts
are not clipped to administrative polygons or normalized by building stock.
"""
import json
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.patches import PathPatch
from matplotlib.path import Path as PlotPath
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
PAPER = '#F2EFE8'
DARK = '#1E2227'


def project(lon, lat):
    return ((np.asarray(lon) + 87.7) * 111.195 * np.cos(np.deg2rad(41.85)),
            (np.asarray(lat) - 41.85) * 111.195)


def rings(filename):
    data = json.loads((ROOT / 'data/geography' / filename).read_text())
    for feature in data['features']:
        geometry = feature['geometry']
        polygons = geometry['coordinates'] if geometry['type'] == 'MultiPolygon' else [geometry['coordinates']]
        for polygon in polygons:
            vertices, codes = [], []
            for i, ring in enumerate(polygon):
                points = np.asarray(ring)
                points = np.column_stack(project(points[:, 0], points[:, 1]))
                signed_area = np.sum(points[:-1, 0]*points[1:, 1] - points[1:, 0]*points[:-1, 1])
                # Opposite winding for holes preserves lake islands accurately.
                if (signed_area > 0) != (i == 0):
                    points = points[::-1]
                vertices.extend(points)
                codes.extend([PlotPath.MOVETO] + [PlotPath.LINETO]*(len(points)-2) + [PlotPath.CLOSEPOLY])
            yield PlotPath(vertices, codes)


def inspection_map(inspections):
    """Return a figure, leaving its title to the shared EDA export routine."""
    fig = plt.figure(figsize=(9, 4.5), facecolor=PAPER)
    ax = fig.add_axes([.015, .12, .79, .75], facecolor='#E5E1D8')
    for ring in rings('lake_michigan.geojson'):
        ax.add_patch(PathPatch(ring, facecolor='#C9DFE4', edgecolor='#8AAEB8', linewidth=.65, zorder=1))
    areas = list(rings('chicago_community_areas.geojson'))
    for ring in areas:
        ax.add_patch(PathPatch(ring, facecolor='#FBF9F4', edgecolor='none', zorder=2))
    coords = inspections.dropna(subset=['longitude', 'latitude'])
    x, y = project(coords.longitude, coords.latitude)
    hb = ax.hexbin(x, y, gridsize=55, bins='log', cmap='Oranges', mincnt=1,
                   linewidths=0, alpha=.85, zorder=3)
    assert int(np.sum(hb.get_array())) == len(coords), 'Map must count every valid inspection once'
    for ring in areas:
        ax.add_patch(PathPatch(ring, facecolor='none', edgecolor='#646A6B', linewidth=.38, alpha=.8, zorder=4))
    ax.set_xlim(*project(np.array([-88.02, -87.36]), 41.85)[0])
    ax.set_ylim(*project(-87.7, np.array([41.62, 42.05]))[1])
    ax.set_aspect('equal')
    ax.set_xticks([]); ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(True); spine.set_color('#C9C3B6'); spine.set_linewidth(.6)
    halo = [pe.withStroke(linewidth=2.5, foreground=PAPER)]
    ax.text(.025, .955, 'ЧИКАГО, США', transform=ax.transAxes, fontsize=11,
            weight='bold', color=DARK, va='top', zorder=6)
    for text, lon, lat, label_lon, label_lat in [
        ("O’Hare", -87.905, 41.978, -87.977, 41.959),
        ('Rogers Park', -87.671, 42.010, -87.615, 42.031),
        ('Austin', -87.763, 41.893, -87.925, 41.886),
        ('Loop · центр', -87.625, 41.883, -87.525, 41.903),
        ('Hyde Park', -87.591, 41.794, -87.515, 41.804),
        ('Beverly', -87.676, 41.718, -87.862, 41.722),
        ('Hegewisch', -87.550, 41.655, -87.485, 41.659),
    ]:
        ax.annotate(text, xy=project(lon, lat), xytext=project(label_lon, label_lat),
                    fontsize=8, color=DARK, va='center', zorder=7, path_effects=halo,
                    arrowprops={'arrowstyle': '-', 'color': '#454B51', 'lw': .65})
    ax.text(*project(-87.49, 41.96), 'озеро\nМИЧИГАН', fontsize=10, color='#47727F',
            ha='center', va='center', linespacing=1.5, zorder=5)
    ax.annotate('С', xy=(.95, .93), xytext=(.95, .79), xycoords='axes fraction',
                ha='center', fontsize=9, color=DARK,
                arrowprops={'arrowstyle': '-|>', 'color': DARK, 'lw': 1}, zorder=6)
    sx, sy = project(-87.975, 41.65)
    ax.plot([sx, sx+5], [sy, sy], color=DARK, lw=2, zorder=6)
    ax.plot([sx, sx, sx+5, sx+5], [sy-.3, sy+.3, sy+.3, sy-.3], color=DARK, lw=.8, zorder=6)
    ax.text(sx+2.5, sy+1, '5 км', ha='center', fontsize=8, color=DARK, zorder=6)
    cax = fig.add_axes([.845, .35, .019, .41])
    cb = fig.colorbar(hb, cax=cax)
    cb.ax.tick_params(labelsize=9)
    cb.outline.set_linewidth(.5)
    fig.text(.828, .81, 'Инспекций\nв ячейке', fontsize=9, color=DARK, linespacing=1.4)
    fig.text(.825, .215, 'Логарифмическая\nшкала', fontsize=8, color='#5B636E', linespacing=1.4)
    fig.text(.19, .055, 'Контуры: 77 районов Чикаго  ·  Подложка: City of Chicago / Natural Earth',
             fontsize=6.8, color='#5B636E')
    return fig


if __name__ == '__main__':
    import pandas as pd
    data = pd.read_parquet(ROOT / 'data/processed/inspections.parquet')
    fig = inspection_map(data)
    fig.suptitle('7. Пространственная концентрация инспекций', fontsize=14, fontweight='bold', x=.03, ha='left')
    for ext in ['png', 'svg']:
        fig.savefig(ROOT / f'reports/figures/eda07_geo.{ext}', dpi=180)
    plt.close(fig)
