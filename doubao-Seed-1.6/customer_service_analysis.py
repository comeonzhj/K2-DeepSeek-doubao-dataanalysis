import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime

# 设置中文字体，确保中文正常显示
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

def analyze_customer_service_data():
    """分析客户服务满意度数据并生成可视化报告"""
    # 读取Excel文件
    excel_file = "客户服务满意度提升数据.xlsx"
    
    try:
        # 使用pandas读取Excel数据
        df = pd.read_excel(excel_file)
        
        # 创建报告目录
        if not os.path.exists("reports"):
            os.makedirs("reports")
            
        # 生成报告文件名（包含时间戳）
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = f"reports/客户服务质量分析报告_{timestamp}.md"
        visualization_dir = f"reports/visualizations_{timestamp}"
        
        # 创建可视化图表保存目录
        if not os.path.exists(visualization_dir):
            os.makedirs(visualization_dir)
        
        # 开始生成Markdown报告
        with open(report_path, "w", encoding="utf-8") as report:
            # 报告标题
            report.write("# 客户服务质量分析报告\n\n")
            report.write(f"**生成时间:** {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}\n\n")
            
            # 1. 数据概览
            report.write("## 1. 数据概览\n\n")
            report.write(f"本次分析共包含 **{len(df)}** 条客户服务记录，涉及 **{len(df.columns)}** 个字段。\n\n")
            
            # 显示数据字段信息
            report.write("### 数据字段说明\n\n")
            report.write("| 字段名 | 数据类型 | 非空值数量 | 空值比例 |\n")
            report.write("|--------|----------|------------|----------|\n")
            
            for column in df.columns:
                dtype = str(df[column].dtype)
                non_null_count = df[column].count()
                null_ratio = f"{(1 - non_null_count/len(df)):.2%}"
                report.write(f"| {column} | {dtype} | {non_null_count} | {null_ratio} |\n")
            
            report.write("\n### 数据样例（前5条记录）\n\n")
            report.write(df.head().to_markdown(index=False))
            report.write("\n\n")
            
            # 2. 描述性统计分析
            report.write("## 2. 描述性统计分析\n\n")
            
            # 识别数值型列进行统计分析
            numeric_columns = df.select_dtypes(include=['int64', 'float64']).columns
            if len(numeric_columns) > 0:
                report.write("### 数值型字段统计摘要\n\n")
                report.write(df[numeric_columns].describe().to_markdown())
                report.write("\n\n")
                
                # 为每个数值型字段生成直方图
                for col in numeric_columns:
                    plt.figure(figsize=(10, 6))
                    sns.histplot(df[col], kde=True)
                    plt.title(f"{col} 分布情况")
                    plt.tight_layout()
                    plot_path = f"{visualization_dir}/{col}_distribution.png"
                    plt.savefig(plot_path)
                    plt.close()
                    
                    report.write(f"#### {col} 分布\n\n")
                    report.write(f"![{col} 分布]({plot_path})\n\n")
            else:
                report.write("数据中未发现数值型字段，无法进行统计分析。\n\n")
            
            # 3. 相关性分析（如果有足够的数值型字段）
            if len(numeric_columns) >= 2:
                report.write("## 3. 相关性分析\n\n")
                
                # 计算相关系数
                corr_matrix = df[numeric_columns].corr()
                
                # 生成热力图
                plt.figure(figsize=(12, 10))
                sns.heatmap(corr_matrix, annot=True, cmap="coolwarm", vmin=-1, vmax=1)
                plt.title("各指标相关性热力图")
                plt.tight_layout()
                corr_plot_path = f"{visualization_dir}/correlation_heatmap.png"
                plt.savefig(corr_plot_path)
                plt.close()
                
                report.write("### 各指标相关性热力图\n\n")
                report.write(f"![相关性热力图]({corr_plot_path})\n\n")
                
                # 找出强相关性
                strong_correlations = []
                for i in range(len(corr_matrix.columns)):
                    for j in range(i):
                        if abs(corr_matrix.iloc[i, j]) > 0.7:  # 强相关阈值
                            strong_correlations.append({
                                '指标1': corr_matrix.columns[i],
                                '指标2': corr_matrix.columns[j],
                                '相关系数': corr_matrix.iloc[i, j]
                            })
                
                if strong_correlations:
                    report.write("### 强相关性指标（相关系数 > 0.7 或 < -0.7）\n\n")
                    report.write("| 指标1 | 指标2 | 相关系数 |\n")
                    report.write("|-------|-------|----------|\n")
                    for corr in strong_correlations:
                        report.write(f"| {corr['指标1']} | {corr['指标2']} | {corr['相关系数']:.4f} |\n")
                    report.write("\n\n")
            
            # 4. 分类变量分析
            report.write("## 4. 分类变量分析\n\n")
            categorical_columns = df.select_dtypes(include=['object', 'category']).columns
            
            if len(categorical_columns) > 0:
                for col in categorical_columns:
                    # 计算各类别占比
                    value_counts = df[col].value_counts(normalize=True)
                    
                    # 生成饼图
                    plt.figure(figsize=(10, 6))
                    value_counts.plot(kind='pie', autopct='%1.1f%%')
                    plt.title(f"{col} 分布占比")
                    plt.ylabel('')  # 移除y轴标签
                    plt.tight_layout()
                    pie_path = f"{visualization_dir}/{col}_pie.png"
                    plt.savefig(pie_path)
                    plt.close()
                    
                    report.write(f"### {col} 分布\n\n")
                    report.write(f"![{col} 饼图]({pie_path})\n\n")
                    
                    # 如果有数值型字段，尝试按分类字段分组比较
                    if len(numeric_columns) > 0:
                        # 选择第一个数值型字段进行分组比较
                        num_col = numeric_columns[0]
                        plt.figure(figsize=(12, 7))
                        sns.boxplot(x=col, y=num_col, data=df)
                        plt.title(f"{col} 与 {num_col} 的关系")
                        plt.xticks(rotation=45)
                        plt.tight_layout()
                        box_path = f"{visualization_dir}/{col}_vs_{num_col}_boxplot.png"
                        plt.savefig(box_path)
                        plt.close()
                        
                        report.write(f"#### {col} 与 {num_col} 的关系\n\n")
                        report.write(f"![{col} 与 {num_col} 箱线图]({box_path})\n\n")
            else:
                report.write("数据中未发现分类字段。\n\n")
            
            # 5. 业务优化建议
            report.write("## 5. 业务优化建议\n\n")
            
            # 基于数据分析生成建议
            suggestions = []
            
            # 如果有满意度相关字段
            satisfaction_cols = [col for col in df.columns if '满意' in col or '评分' in col or 'score' in col.lower()]
            if satisfaction_cols:
                for col in satisfaction_cols:
                    avg_score = df[col].mean()
                    suggestions.append(f"- **提升{col}**：当前平均{col}为{avg_score:.2f}，建议针对评分较低的服务环节进行专项改进。")
            
            # 如果有投诉或问题相关字段
            complaint_cols = [col for col in df.columns if '投诉' in col or '问题' in col or '抱怨' in col or 'complain' in col.lower()]
            if complaint_cols:
                suggestions.append(f"- **减少客户投诉**：分析{', '.join(complaint_cols)}字段，识别主要投诉类型并制定针对性解决方案。")
            
            # 如果有时间相关字段
            time_cols = [col for col in df.columns if '时间' in col or '日期' in col or 'time' in col.lower() or 'date' in col.lower()]
            if time_cols:
                suggestions.append(f"- **优化服务时效**：分析{', '.join(time_cols)}字段，找出服务高峰期和响应延迟时段，合理调配人力资源。")
            
            # 如果有员工或部门相关字段
            staff_cols = [col for col in df.columns if '员工' in col or '部门' in col or 'staff' in col.lower() or 'department' in col.lower()]
            if staff_cols and satisfaction_cols:
                suggestions.append(f"- **加强员工培训**：结合{', '.join(staff_cols)}和{', '.join(satisfaction_cols)}字段，识别表现优异的员工/部门经验并推广，对表现不佳的进行针对性培训。")
            
            # 如果没有足够数据生成针对性建议，提供通用建议
            if not suggestions:
                suggestions = [
                    "- **提升响应速度**：优化客户服务响应流程，减少客户等待时间。",
                    "- **加强员工培训**：提高客服人员沟通技巧和问题解决能力。",
                    "- **完善反馈机制**：建立更便捷的客户反馈渠道，及时了解客户需求。",
                    "- **个性化服务**：根据客户历史数据，提供更具针对性的服务方案。",
                    "- **优化服务流程**：简化客户服务流程，减少不必要的环节和等待。"
                ]
            
            # 将建议写入报告
            for s in suggestions:
                report.write(f"{s}\n\n")
            
            # 6. 总结
            report.write("## 6. 总结\n\n")
            report.write("本报告通过对客户服务数据的全面分析，揭示了当前服务质量的主要特点和潜在问题。\n")
            report.write("建议根据上述分析结果和优化建议，制定具体的改进计划并跟踪实施效果。\n")
            report.write("定期进行类似数据分析，持续监控服务质量变化趋势。\n\n")
            
            # 报告结束
            report.write("---\n")
            report.write("*报告生成工具：客户服务数据分析脚本 v1.0*")
        
        print(f"分析完成！报告已保存至：{report_path}")
        print(f"可视化图表已保存至：{visualization_dir}")
        return report_path
        
    except Exception as e:
        print(f"分析过程中出错：{str(e)}")
        return None

if __name__ == "__main__":
    analyze_customer_service_data()
