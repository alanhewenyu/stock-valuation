# stock-valuation

## DCF估值介绍
预测未来自由现金流（FCFF）并折现，是公司估值的常见做法，也叫绝对估值法。这种方法计算的公司价值也叫内在价值（intrinsic value）。

价值投资领域，常用内在价值与股票的市场价值做比较，以判断公司价值是低估还是高估。

DCF模型通常需要对公司的财务报表做详细预测，这一工作量往往是繁琐而且巨大的。为了让更多人能够应用DCF模型估算股票价值，本项目对模型做了简化处理，选取核心变量，减少参数输入。

估值是一门艺术。巴菲特说，“我宁愿模糊的正确，也不愿精确的错误”。通过模糊估值，我们可以更好理解股票估值背后的关键驱动因素，并通过敏感性分析找到股票的安全垫。

## 项目做什么

本项目旨在基于内置的DCF估值模型计算股票的内在价值，并根据核心变量做估值敏感性分析。由于大量的基础财务数据通过api抓取，你可通过运行python script快速计算出股票的内在价值。

## 如何使用项目

### 万得API

为了使用本项目，首先要有万得金融终端的使用权限。如果没有万得账户，请忽略本项目。

### DCF估值模版

本项目按照上述DCF模型搭建了计算表（stock valuation template）,请下载计算表，并放入要计算存放估值结果的文件路径。

计算表中的大多数数据都通过万得API自动抓取并输入，仅有少数参数需要通过电脑终端引导输入。

程序运行结束后，计算表数据更新完毕，即可算出股票的合理价格。计算表会基于两个核心参数--收入增长率和经营利润率，对估值结果进一步做敏感性分析。

### 具体运行步骤如下：

本项目在电脑终端运行：

(1) 下载stock.py和stock valuation template.xlsx

(2) 修改stock.py中的file route。file route为存放stock valuation template.xlsx的路径。

(3) 运行 python stock.py

(4) 根据提示输入股票代码，生成股票过去5年关键财务指标，以辅助判断未来预测指标

(5) 按照终端提示，输入估值参数

(6) 估值计算表自动更新

### 需要手工输入的参数可分为四大项，12个：

#### 收入增长：

Base year for valuation（用于预测的基准年度）

Revenue growth for year 1 forecast（第一个预测年度的收入增长率）

Compound annual revenue growth rate (FY2-5)（2至5年度的收入增长率）

Revenue growth for terminal year (equal to risk-free rate)（永续期收入增长率）

#### 经营利润率：

Terminal target EBIT margin（到永续期为止的目标EBIT利润率）

Years of convergence for target margin（达到目标EBIT利润率所需的年份数）

#### 资本开支：

Revenue to capital ratio (next 2 years)（1至2年度收入/投入资本比率）

Revenue to capital ratio (FY3-5)（3至5年度收入/投入资本比率）

Revenue to capital ratio (FY5-10)（5至10年度收入/投入资本比率）

#### 其他：

Terminal WACC（到永续期的加权平均资本成本）

Return on new invested capital for the long term（判断选项，长期资本投入收益率是否等于WACC？）

Effective tax rate （有效税率）


## 如需了解更多关于公司估值方面的知识，欢迎关注本人微信公众号《见山笔记》。
