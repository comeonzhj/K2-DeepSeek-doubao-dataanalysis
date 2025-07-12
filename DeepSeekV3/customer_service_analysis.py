import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# 读取数据
df = pd.read_excel('客户服务满意度提升数据.xlsx')

# 1. 基本数据概况
total_cases = len(df)
resolved_rate = df['是否解决'].mean()
avg_satisfaction = df['满意度评分'].mean()
avg_resolution_time = df['解决时长_分钟'].mean()

# 2. 服务渠道分析
channel_dist = df['服务渠道'].value_counts(normalize=True)
channel_satisfaction = df.groupby('服务渠道')['满意度评分'].mean().sort_values()

# 3. 问题类型分析
issue_dist = df['问题类型'].value_counts(normalize=True)
issue_satisfaction = df.groupby('问题类型')['满意度评分'].mean().sort_values()

# 4. 解决时长与满意度关系
time_satisfaction_corr = df['解决时长_分钟'].corr(df['满意度评分'])

# 生成Markdown报告
with open('customer_service_report.md', 'w') as f:
    f.write("# 客户服务质量分析报告\n\n")
    f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    
    # 基本概况
    f.write("## 1. 基本概况\n")
    f.write(f"- 总服务案例数: {total_cases}\n")
    f.write(f"- 问题解决率: {resolved_rate:.1%}\n")
    f.write(f"- 平均满意度评分: {avg_satisfaction:.1f}/5\n")
    f.write(f"- 平均解决时长: {avg_resolution_time:.1f}分钟\n\n")
    
    # 服务渠道分析
    f.write("## 2. 服务渠道分析\n")
    f.write("### 各渠道使用分布\n")
    f.write(channel_dist.to_markdown())
    f.write("\n\n### 各渠道满意度对比\n")
    f.write(channel_satisfaction.to_markdown())
    f.write("\n\n![渠道满意度](channel_satisfaction.png)\n\n")
    
    # 问题类型分析
    f.write("## 3. 问题类型分析\n")
    f.write("### 各类问题分布\n")
    f.write(issue_dist.to_markdown())
    f.write("\n\n### 各类问题满意度对比\n")
    f.write(issue_satisfaction.to_markdown())
    f.write("\n\n![问题类型满意度](issue_satisfaction.png)\n\n")
    
    # 解决时长分析
    f.write("## 4. 解决时长分析\n")
    f.write(f"- 解决时长与满意度相关性: {time_satisfaction_corr:.2f}\n")
    f.write("\n![解决时长分布](resolution_time_dist.png)\n")
    f.write("![解决时长vs满意度](time_vs_satisfaction.png)\n\n")
    
    # 优化建议
    f.write("## 5. 优化建议\n")
    f.write("1. **渠道优化**: 满意度较低的渠道需要改进服务质量或流程\n")
    f.write("2. **问题类型聚焦**: 对高频低满意度问题类型优先优化\n")
    f.write("3. **效率提升**: 缩短解决时长，特别是对满意度影响大的问题\n")
    f.write("4. **培训重点**: 针对低满意度渠道和问题类型加强客服培训\n")
    f.write("5. **流程优化**: 分析解决时长过长的案例，优化处理流程\n")

# 生成可视化图表
plt.figure(figsize=(10,6))
channel_satisfaction.plot(kind='bar')
plt.title('各服务渠道平均满意度')
plt.ylabel('满意度评分(1-5)')
plt.savefig('channel_satisfaction.png')
plt.close()

plt.figure(figsize=(10,6))
issue_satisfaction.plot(kind='bar')
plt.title('各类问题平均满意度')
plt.ylabel('满意度评分(1-5)')
plt.savefig('issue_satisfaction.png')
plt.close()

plt.figure(figsize=(10,6))
sns.histplot(df['解决时长_分钟'], bins=20)
plt.title('解决时长分布')
plt.xlabel('分钟')
plt.savefig('resolution_time_dist.png')
plt.close()

plt.figure(figsize=(10,6))
sns.scatterplot(x='解决时长_分钟', y='满意度评分', data=df)
plt.title('解决时长与满意度关系')
plt.savefig('time_vs_satisfaction.png')
plt.close()
