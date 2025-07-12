import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# 读取数据
df = pd.read_excel('客户服务满意度提升数据.xlsx')

# 数据预处理
df['服务日期'] = pd.to_datetime(df['服务日期'])
df['月份'] = df['服务日期'].dt.month
df['星期'] = df['服务日期'].dt.day_name()

# 基本统计信息
print("=== 数据概览 ===")
print(f"总记录数: {len(df)}")
print(f"时间范围: {df['服务日期'].min()} 到 {df['服务日期'].max()}")
print(f"客户数量: {df['客户ID'].nunique()}")

# 满意度分布
satisfaction_dist = df['满意度评分'].value_counts().sort_index()
print("\n=== 满意度评分分布 ===")
print(satisfaction_dist)

# 服务渠道分析
channel_analysis = df.groupby('服务渠道').agg({
    '服务ID': 'count',
    '满意度评分': ['mean', 'std'],
    '解决时长_分钟': 'mean',
    '是否解决': 'mean'
}).round(2)
channel_analysis.columns = ['服务数量', '平均满意度', '满意度标准差', '平均解决时长', '解决率']
print("\n=== 各服务渠道表现 ===")
print(channel_analysis)

# 问题类型分析
problem_analysis = df.groupby('问题类型').agg({
    '服务ID': 'count',
    '满意度评分': 'mean',
    '解决时长_分钟': 'mean',
    '是否解决': 'mean'
}).round(2)
problem_analysis.columns = ['问题数量', '平均满意度', '平均解决时长', '解决率']
problem_analysis = problem_analysis.sort_values('问题数量', ascending=False)
print("\n=== 各问题类型表现 ===")
print(problem_analysis)

# 创建可视化图表
fig, axes = plt.subplots(2, 3, figsize=(18, 12))
fig.suptitle('客户服务满意度分析报告', fontsize=16, fontweight='bold')

# 1. 满意度评分分布
axes[0,0].pie(satisfaction_dist.values, labels=[f'{i}分' for i in satisfaction_dist.index], 
              autopct='%1.1f%%', colors=['#ff9999','#ffcc99','#ffff99','#99ff99','#66b3ff'])
axes[0,0].set_title('客户满意度评分分布')

# 2. 各服务渠道的满意度
channel_satisfaction = df.groupby('服务渠道')['满意度评分'].mean().sort_values(ascending=False)
axes[0,1].bar(channel_satisfaction.index, channel_satisfaction.values, color='skyblue')
axes[0,1].set_title('各服务渠道平均满意度')
axes[0,1].set_ylabel('平均满意度评分')
axes[0,1].tick_params(axis='x', rotation=45)

# 3. 解决时长与满意度关系
axes[0,2].scatter(df['解决时长_分钟'], df['满意度评分'], alpha=0.5, color='green')
axes[0,2].set_xlabel('解决时长（分钟）')
axes[0,2].set_ylabel('满意度评分')
axes[0,2].set_title('解决时长与满意度关系')

# 4. 各问题类型数量
axes[1,0].bar(problem_analysis.index, problem_analysis['问题数量'], color='lightcoral')
axes[1,0].set_title('各问题类型数量')
axes[1,0].set_ylabel('问题数量')
axes[1,0].tick_params(axis='x', rotation=45)

# 5. 月度满意度趋势
monthly_satisfaction = df.groupby('月份')['满意度评分'].mean()
axes[1,1].plot(monthly_satisfaction.index, monthly_satisfaction.values, marker='o', linewidth=2, markersize=8)
axes[1,1].set_xlabel('月份')
axes[1,1].set_ylabel('平均满意度评分')
axes[1,1].set_title('月度满意度趋势')
axes[1,1].set_xticks(range(1, 13))

# 6. 解决率对比
resolution_rate = df.groupby('服务渠道')['是否解决'].mean().sort_values(ascending=False)
axes[1,2].bar(resolution_rate.index, resolution_rate.values * 100, color='gold')
axes[1,2].set_title('各渠道问题解决率')
axes[1,2].set_ylabel('解决率 (%)')
axes[1,2].tick_params(axis='x', rotation=45)

plt.tight_layout()
plt.savefig('客户服务满意度分析图表.png', dpi=300, bbox_inches='tight')
plt.show()

# 深度分析
print("\n=== 深度分析 ===")

# 低满意度分析（1-2分）
low_satisfaction = df[df['满意度评分'] <= 2]
print(f"\n低满意度服务（1-2分）占比: {len(low_satisfaction)/len(df)*100:.1f}%")
print("低满意度服务的主要特征:")
print(f"- 主要问题类型: {low_satisfaction['问题类型'].value_counts().head(3).to_dict()}")
print(f"- 主要服务渠道: {low_satisfaction['服务渠道'].value_counts().head(3).to_dict()}")
print(f"- 平均解决时长: {low_satisfaction['解决时长_分钟'].mean():.1f}分钟")

# 高满意度分析（4-5分）
high_satisfaction = df[df['满意度评分'] >= 4]
print(f"\n高满意度服务（4-5分）占比: {len(high_satisfaction)/len(df)*100:.1f}%")
print("高满意度服务的主要特征:")
print(f"- 主要问题类型: {high_satisfaction['问题类型'].value_counts().head(3).to_dict()}")
print(f"- 主要服务渠道: {high_satisfaction['服务渠道'].value_counts().head(3).to_dict()}")
print(f"- 平均解决时长: {high_satisfaction['解决时长_分钟'].mean():.1f}分钟")

# 相关性分析
correlation = df[['解决时长_分钟', '满意度评分', '是否解决']].corr()
print("\n=== 相关性分析 ===")
print(correlation)

# 保存分析结果到文件
with open('客户服务满意度分析报告.md', 'w', encoding='utf-8') as f:
    f.write("# 客户服务满意度提升数据分析报告\n\n")
    
    f.write("## 1. 数据概览\n")
    f.write(f"- **总服务记录数**: {len(df)}\n")
    f.write(f"- **时间范围**: {df['服务日期'].min().strftime('%Y-%m-%d')} 至 {df['服务日期'].max().strftime('%Y-%m-%d')}\n")
    f.write(f"- **服务客户数**: {df['客户ID'].nunique()}\n")
    f.write(f"- **平均满意度**: {df['满意度评分'].mean():.2f}\n")
    f.write(f"- **整体解决率**: {df['是否解决'].mean()*100:.1f}%\n\n")
    
    f.write("## 2. 满意度评分分布\n")
    f.write("| 满意度评分 | 数量 | 占比 |\n")
    f.write("|------------|------|------|\n")
    for score, count in satisfaction_dist.items():
        percentage = count/len(df)*100
        f.write(f"| {score}分 | {count} | {percentage:.1f}% |\n")
    
    f.write("\n## 3. 各服务渠道表现分析\n")
    f.write("| 服务渠道 | 服务数量 | 平均满意度 | 平均解决时长(分钟) | 解决率 |\n")
    f.write("|----------|----------|------------|-------------------|--------|\n")
    for channel in channel_analysis.index:
        row = channel_analysis.loc[channel]
        f.write(f"| {channel} | {int(row['服务数量'])} | {row['平均满意度']:.2f} | {row['平均解决时长']:.1f} | {row['解决率']*100:.1f}% |\n")
    
    f.write("\n## 4. 问题类型分析\n")
    f.write("| 问题类型 | 问题数量 | 平均满意度 | 平均解决时长(分钟) | 解决率 |\n")
    f.write("|----------|----------|------------|-------------------|--------|\n")
    for problem in problem_analysis.index:
        row = problem_analysis.loc[problem]
        f.write(f"| {problem} | {int(row['问题数量'])} | {row['平均满意度']:.2f} | {row['平均解决时长']:.1f} | {row['解决率']*100:.1f}% |\n")

print("\n分析完成！图表已保存为 '客户服务满意度分析图表.png'")
print("详细报告已保存为 '客户服务满意度分析报告.md'")
