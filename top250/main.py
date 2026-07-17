import sys
import os
from utils import IMDBCrawler, IMDBAnalyzer, ExcelDashboardGenerator

def run_crawler(excel_path="imdb_dashboard.xlsx"):
    print("start extracting data ...")
    crawler = IMDBCrawler(excel_path=excel_path)
    crawler.run()

def run_analysis(excel_path="imdb_dashboard.xlsx", output_dashboard="imdb_dashboard_analyzed.xlsx"):
    print("start analys...")
    analyzer = IMDBAnalyzer(excel_path=excel_path)
    analyzer.generate_report()

    generator = ExcelDashboardGenerator(analyzer.df, output_path=output_dashboard)
    generator.create()

if __name__ == "__main__":
    if not os.path.exists("imdb_dashboard.xlsx"):
        print("⚠️ فایل imdb_dashboard.xlsx یافت نشد. لطفاً ابتدا آن را در مسیر قرار دهید.")
        sys.exit(1)

    run_analysis("imdb_dashboard.xlsx", "imdb_dashboard_analyzed.xlsx")
    # run_crawler("imdb_dashboard.xlsx")
