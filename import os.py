import os
import logging
import sys
import argparse
from pathlib import Path
import pandas as pd

# 配置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def modify_excel(input_dir, output_dir, input_file_name, output_file_name):
    # 构建文件路径
    file_path = Path(input_dir) / input_file_name
    output_path = Path(output_dir) / output_file_name

    # 检查文件是否存在
    if not file_path.exists():
        raise FileNotFoundError(f"文件 {file_path} 不存在")

    try:
        # 读取整个工作簿
        with pd.ExcelFile(file_path) as xls:
            sheet_names = xls.sheet_names
            
            # 处理所有工作表
            processed_dfs = {}
            for sheet_name in sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    
                    # 原数据处理逻辑
                    def process_blank_plate(spec):
                        if not spec or not isinstance(spec, str):
                            return spec
                        parts = spec.split()
                        for i, part in enumerate(parts):
                            if part == 'FF':
                                parts[i] = '-1-'
                            elif part == 'MFM':
                                parts[i] = '-2-'
                            elif part.startswith('PN'):
                                try:
                                    pn = int(part[2:]) / 10
                                    parts[i] = f'-{pn:.1f}'
                                except ValueError:
                                    logging.warning(f"无法解析 PN 值: {part}")
                                    continue
                        return ' '.join(parts)

                    if '盲板规格型号' in df.columns:
                        df['盲板规格型号'] = df['盲板规格型号'].astype(str).apply(process_blank_plate)
                    
                    if '垫片型号' in df.columns:
                        df['垫片型号'] = df['垫片型号'].str.strip().str.replace('缠绕', '', regex=False)
                    
                    processed_dfs[sheet_name] = df
                    logging.info(f"成功处理工作表: {sheet_name}")
                except Exception as e:
                    logging.error(f"处理工作表 {sheet_name} 时出错: {e}")
                    continue

        # 确保输出目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 保存新文件
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in processed_dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            logging.info(f"文件已成功保存到 {output_path}")
        except Exception as e:
            logging.error(f"保存文件时出错: {e}")
            raise

    except Exception as e:
        logging.error(f"读取文件时出错: {e}")
        raise

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="修改 Excel 文件")
    # 修改默认输入目录
    parser.add_argument('--input_dir', type=str, default=r'D:\VS Code EX\aaa', help='输入目录')
    parser.add_argument('--output_dir', type=str, default='.', help='输出目录')
    parser.add_argument('--input_file_name', type=str, default='气分装置盲板台账.xlsx', help='输入文件名')
    parser.add_argument('--output_file_name', type=str, default='气分装置盲板台账新.xlsx', help='输出文件名')
    args = parser.parse_args()

    modify_excel(args.input_dir, args.output_dir, args.input_file_name, args.output_file_name)
    # 删除重复调用
    # modify_excel(args.input_dir, args.output_dir, args.input_file_name, args.output_file_name)