import akshare as ak

sh = ak.stock_info_sh_name_code("主板A股")
print(sh.head(5))

sz = ak.stock_info_sz_name_code("A股列表")
print(sz.head(5))

fund = ak.fund_etf_category_sina("ETF基金")
print(fund.head(5))