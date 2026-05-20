import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# ─────────────────────────────────────────────
# 微软雅黑字体（跨平台）
# ─────────────────────────────────────────────
font_path_list = [
    'C:/Windows/Fonts/msyh.ttc',  # Windows
    '/System/Volumes/Data/Applications/Microsoft Excel.app/Contents/Resources/DFonts/msyh.ttc',  # Mac
    '/System/Volumes/Data/Users/moxiaoxian/Library/Caches/camoufox/Camoufox.app/Contents/Resources/fonts/msyh.ttc',
]

font_loaded = False

for fp in font_path_list:
    try:
        fm.fontManager.addfont(fp)
        font_loaded = True
        break
    except Exception:
        continue

if font_loaded:
    plt.rcParams['font.family'] = 'Microsoft YaHei'
else:
    plt.rcParams['font.family'] = 'sans-serif'

plt.rcParams['axes.unicode_minus'] = False

# ─────────────────────────────────────────────
# 读取 Excel
# 自动读取前两个 Sheet
# ─────────────────────────────────────────────
excel_path = './p459.xlsx'

excel_file = pd.ExcelFile(excel_path)

# 获取前两个 sheet 名称
sheet1_name = excel_file.sheet_names[0]
sheet2_name = excel_file.sheet_names[1]

# 读取数据
s1 = pd.read_excel(excel_file, sheet_name=sheet1_name)
s2 = pd.read_excel(excel_file, sheet_name=sheet2_name)

# 中文列名
cols = ['外径', '内径', '线径']

# 英文列名
cols_en = [
    'Outer Diameter',
    'Inner Diameter',
    'Wire Diameter'
]

# ─────────────────────────────────────────────
# 配色
# ─────────────────────────────────────────────
# Sheet1
fc1 = '#FFD4D4'   # 浅红
ec1 = '#E74C3C'   # 红

# Sheet2
fc2 = '#D4E8FF'   # 浅蓝
ec2 = '#3498DB'   # 蓝

# ─────────────────────────────────────────────
# 全局样式
# ─────────────────────────────────────────────
plt.rcParams.update({
    'font.size': 9,
    'axes.linewidth': 0.6,
    'axes.edgecolor': '#444444',
    'xtick.color': '#444444',
    'ytick.color': '#444444',
    'figure.facecolor': 'white',
    'axes.facecolor': '#FAFBFC',

    'axes.grid.axis': 'y',
    'grid.color': '#E0E0E0',
    'grid.alpha': 0.4,
    'grid.linestyle': '--',
    'grid.linewidth': 0.4,
})

# ─────────────────────────────────────────────
# 绘图
# ─────────────────────────────────────────────
for idx, col in enumerate(cols):

    # 去除空值
    d1 = s1[col].dropna().values
    d2 = s2[col].dropna().values

    # 创建图像
    fig, ax = plt.subplots(
        1,
        1,
        figsize=(120 / 25.4, 110 / 25.4),  # mm 转 inch
        constrained_layout=True
    )

    # 箱线图
    bp = ax.boxplot(
        [d1, d2],
        positions=[1, 2],
        widths=0.5,
        patch_artist=True,
        showmeans=True,

        meanprops=dict(
            marker='D',
            markerfacecolor='#333333',
            markeredgecolor='#333333',
            markersize=6
        ),

        medianprops=dict(
            color='#222222',
            linewidth=1.6,
            solid_capstyle='round'
        ),

        whiskerprops=dict(
            color='#555555',
            linewidth=0.8
        ),

        capprops=dict(
            color='#555555',
            linewidth=0.8
        ),

        flierprops=dict(
            marker='o',
            markersize=7,
            markerfacecolor='none',
            markeredgewidth=1.8,
            linestyle='none'
        ),

        boxprops=dict(
            linewidth=1.0,
            zorder=4
        ),
    )

    # 设置箱体颜色
    bp['boxes'][0].set(
        facecolor=fc1,
        edgecolor=ec1
    )

    bp['boxes'][1].set(
        facecolor=fc2,
        edgecolor=ec2
    )

    # 异常值颜色
    bp['fliers'][0].set(
        markeredgecolor=ec1,
        markerfacecolor=fc1
    )

    bp['fliers'][1].set(
        markeredgecolor=ec2,
        markerfacecolor=fc2
    )

    # ─────────────────────────────────────────
    # 均值线
    # ─────────────────────────────────────────
    mean1 = np.mean(d1)
    mean2 = np.mean(d2)

    for pos, mean_val in [(1, mean1), (2, mean2)]:
        ax.hlines(
            y=mean_val,
            xmin=pos - 0.25,
            xmax=pos + 0.25,
            color='#333333',
            linewidth=2.0,
            zorder=7
        )

    # ─────────────────────────────────────────
    # 散点（Jitter）
    # ─────────────────────────────────────────
    rng = np.random.default_rng(42 + idx)

    ax.scatter(
        1 + rng.uniform(-0.12, 0.12, len(d1)),
        d1,
        alpha=0.30,
        s=12,
        color=ec1,
        edgecolors='none',
        zorder=3
    )

    ax.scatter(
        2 + rng.uniform(-0.12, 0.12, len(d2)),
        d2,
        alpha=0.30,
        s=12,
        color=ec2,
        edgecolors='none',
        zorder=3
    )

    # ─────────────────────────────────────────
    # 标题与坐标轴
    # ─────────────────────────────────────────
    ax.set_title(
        f'{col} ({cols_en[idx]})',
        fontsize=11,
        fontweight='bold',
        pad=8
    )

    ax.set_xticks([1, 2])

    # 自动使用 sheet 名称
    ax.set_xticklabels(
        [sheet1_name, sheet2_name],
        fontsize=9,
        fontweight='medium'
    )

    ax.set_ylabel(
        'Measurement (mm)',
        fontsize=9
    )

    # 网格
    ax.yaxis.grid(
        True,
        linestyle='--',
        alpha=0.25,
        color='#CCCCCC'
    )

    # 去除顶部与右侧边框
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)

    # 左下角单位 mm
    ax.annotate(
        'mm',
        xy=(0, 0),
        xycoords='axes fraction',

        xytext=(4, -22),
        textcoords='offset points',

        ha='right',
        va='top',

        fontsize=8,
        color='#444444',
        fontstyle='italic'
    )

    # ─────────────────────────────────────────
    # 输出文件
    # ─────────────────────────────────────────
    output = f'./{col}_{cols_en[idx].replace(" ", "_")}.png'

    fig.savefig(
        output,
        dpi=300,
        bbox_inches='tight',
        facecolor='white',
        edgecolor='none'
    )

    plt.close(fig)

    print(f'Saved: {output}')

print('\nAll done!')