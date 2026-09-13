**Coca-Cola Sales Analysis Using Excel
**
📊 **Project Overview**

This project analyzes Coca-Cola sales data across U.S. retailers,
regions, states, cities, beverage brands, and months using Microsoft
Excel.

**Business question:** How can Coca-Cola optimize sales and
profitability across U.S. retailers?

The analysis examines revenue, sales volume, operating profit, operating
margin, pricing, geographic performance, retailer performance, brand
performance, and monthly trends, and converts the results into
actionable business recommendations.

🎯 **Business Objectives**

Understand sales variation by retailer, region, state, brand, and
time.

Identify the strongest and weakest retailers and markets.

Compare sales volume, revenue, operating profit, and operating
margin.

Examine the relationship between price per unit and units sold.

Identify monthly and seasonal sales patterns.

Build an Excel workflow using formulas, Pivot Tables, charts,
conditional formatting, and slicers.

Translate findings into business recommendations.

🗂️ **Dataset**

The supplied dataset contains 3,888 records covering 2 January
2021 to 25 December 2021.

**Main dimensions**

Retailer

Retailer ID

Invoice Date

Region

State

City

Beverage Brand

Month

**Main measures**

Price per Unit

Units Sold

Total Sales

Operating Profit

Operating Margin

Metric                                 Value

Records                                3,888
Retailers                                  4
Regions                                    5
States                                    50
Cities                                    52
Beverage Brands                            6
Time period                             2021
Missing values identified                  0
Exact duplicate records identified         0

🧹 **Data Cleaning & Preparation
**
The project includes the following preparation steps:

Review headers and remove blank/non-data rows.

Check duplicate records using Data → Remove Duplicates.

Check critical fields for missing values using COUNTBLANK.

Verify Invoice Date is a valid Excel date.

Verify price, units, sales, and profit fields are numeric.

Check categorical fields for consistency.

Review calculated fields for sales, profit, and month.

Use the Month field for time-based aggregation.

Convert the data range into an Excel Table using Ctrl + T where
appropriate.

Example missing-value check:

=COUNTBLANK(Data!D6:D3893)

The supplied dataset returned 0 blanks for the checked critical
fields.

Calculated fields

=Price_per_Unit*Units_Sold

=Total_Sales*Operating_Margin

=MONTH(Invoice_Date)

📌 **Key KPIs
**
KPI                                              Result

Total Sales                          $8,684,027.50
Total Units Sold                         17,148,250
Operating Profit                     $3,173,631.88
Weighted Operating Margin                     36.5%
Average Price per Unit                       $0.49
Price vs Units Correlation                    0.247
Price vs Total Sales Correlation              0.654

Example formulas

=SUM(Data!K6:K3893)

=SUM(Data!J6:J3893)

=SUM(Data!L6:L3893)

=SUM(Data!L6:L3893)/SUM(Data!K6:K3893)

=AVERAGE(Data!I6:I3893)

=CORREL(Data!I6:I3893,Data!J6:J3893)

=CORREL(Data!I6:I3893,Data!K6:K3893)

🏪 **Retailer Analysis**

Retailers were compared using Total Sales, Units Sold, Operating Profit,
and Operating Margin.

Retailer            Total Sales   Operating Profit   Avg. Margin

Sodapop      $4,403,630.00     $1,647,936.38         36.4%
FizzySip     $2,584,450.00       $835,759.38         32.9%
BevCo        $1,291,535.00       $542,046.13     40.6%
DreamCo        $404,412.50       $147,890.00         38.3%

Insight: Sodapop is the largest retailer by sales and operating
profit, contributing approximately 50.7% of total sales. BevCo has the
highest retailer margin at approximately 40.6%.

🥤 **Brand Analysis**

Brand                   Total Sales   Operating Profit   Avg. Margin

Coca-Cola        $2,015,890.00       $793,197.25         38.7%
Dasani Water     $1,725,837.50       $655,700.00         38.5%
Diet Coke            $1,481,425.00       $494,317.50         34.2%
Sprite               $1,235,587.50       $430,519.63         34.9%
Powerade             $1,193,637.50       $427,071.88         35.3%
Fanta                $1,031,650.00       $372,825.63         36.2%

Insight: Coca-Cola is the leading brand by both sales and operating
profit, contributing approximately 23.2% of total sales.

🌎 **Regional Analysis**

Region           Total Sales   Operating Profit   Avg. Margin

West      $2,835,100.00       $923,014.38         33.1%
Northeast     $1,788,347.50       $629,546.00         34.5%
Southeast     $1,618,935.00       $641,553.88         38.8%
South         $1,291,535.00       $542,046.13     40.6%
Midwest       $1,150,110.00       $437,471.50         37.3%

Insight: West is the largest revenue-generating region and has the
highest operating profit, while South has the strongest operating
margin.

🗺️ **State-Level Analysis**

Selected findings:

State               Total Sales   Operating Profit   Avg. Margin

New York       $582,675.00       $203,082.50         34.0%
California     $582,400.00       $183,895.62         33.0%
Florida        $561,850.00       $214,731.25         37.0%
Nebraska        $54,380.00        $21,991.50         39.0%

Nebraska has the lowest sales among the states highlighted, but its
margin is relatively strong. This suggests that low sales may be related
to demand, market size, or distribution coverage rather than poor
profitability.

📅 **Monthly & Seasonal Analysis**

Key findings:

Highest sales month: December --- approximately $1.046M

Second-highest: July --- approximately $1.041M

Lowest sales month: March --- approximately $0.484M

Sales increase substantially from May through July, soften during
September and October, and rebound during November and December.

Business implication: Coca-Cola can use these patterns for inventory
planning, distribution planning, and promotional timing.

💰 **Pricing Analysis**

A scatter plot was used to examine the relationship between Price per
Unit and Units Sold.

The correlation is approximately 0.247, indicating a modest positive
association.

This does not mean that higher prices cause higher sales. Other factors
may influence volume, including:

Brand

Retailer

Region

Product mix

Distribution

Promotions

Market demand

The price-versus-total-sales correlation is approximately 0.654,
which is a stronger association, but it should also not be interpreted
as proof of causation.

📊 **Visualizations**

Recommended Excel visuals:

Clustered column chart: Retailer vs Total Sales.

Combination chart: Monthly Sales columns + Operating Profit
line.

100% stacked column chart: Monthly sales composition by brand.

State margin heatmap: Conditional Formatting on Operating
Margin.

Scatter plot with trendline: Price per Unit vs Units Sold.

Pivot Tables + Slicers: Interactive analysis by Retailer,
Region, Brand, and Month.

🔄 **Pivot Table Structure**

Retailer Performance

Rows: Retailer

Values: Sum of Total Sales, Sum of Units Sold, Sum of Operating
Profit, Average Operating Margin.

Regional Performance

Rows: Region

Values: Sum of Total Sales, Sum of Units Sold, Sum of Operating
Profit, Average Operating Margin.

Brand Performance

Rows: Beverage Brand

Values: Sum of Total Sales, Sum of Units Sold, Sum of Operating
Profit, Average Operating Margin.

Monthly Trend

Rows: Month

Values: Sum of Total Sales, Sum of Operating Profit, Average
Operating Margin.

State Performance

Rows: State

Values: Sum of Total Sales, Sum of Operating Profit, Average
Operating Margin.

💡 **Key Business Insights**

1. Retailer concentration

Sodapop generates more than half of total sales.

Action: Protect distribution, product availability, and retailer
relationships while studying the drivers of its performance.

2. Revenue vs profitability

West is the largest region by sales, while South has the highest margin.

Action: Study South's product mix, pricing, and operating practices
for potentially transferable practices.

3. Brand leadership

Coca-Cola is the leading brand by sales and operating profit.

Action: Maintain strong availability and evaluate targeted
promotional opportunities.

4. Seasonal demand

July and December are the strongest sales months, while March is the
weakest.

Action: Plan inventory and promotional activity ahead of seasonal
peaks.

5. Pricing is not the whole story

The price-to-units correlation is approximately 0.247.

Action: Analyze retailer, brand, geographic, distribution, and
promotional factors alongside pricing.

6. Low sales does not always mean low profitability

Nebraska has low sales but a relatively strong margin.

Action: Investigate market coverage and demand before reducing
investment.

🚀 **Recommendations**

Protect high-value retailer relationships.

Maintain availability of leading brands, especially Coca-Cola.

Use high-margin regions as benchmarks for operational and pricing
practices.

Plan inventory around seasonal peaks, particularly July and
December.

Investigate underperforming states and retailers at a more granular
level.

Avoid relying on price alone to explain sales volume.

Evaluate sales and profitability together when making market
decisions.

🧰 **Excel Skills Demonstrated**

Data cleaning and validation

Duplicate and missing-value checks

Date handling

SUM

AVERAGE

COUNTBLANK

MONTH

YEAR

TEXT

TRIM

CORREL

Pivot Tables

Pivot Charts

Slicers

Conditional Formatting

Bar and Column Charts

Line Charts

Combination Charts

Scatter Plots

KPI analysis

Trend and profitability analysis

📁 **Workbook Structure**

Sheet                               Purpose

Final Solution                  Completed project plan, actual
results, formulas, insights,
visualization plan, and
presentation storyline

Data                            Coca-Cola transaction-level dataset

Beginner                        Beginner-level Excel analysis

Intermediate                    Intermediate-level analysis

Advanced                        Advanced analysis

Addtional Analysis              Additional analytical work

🎥 Presentation Structure

The project can be presented in approximately 8 minutes:

Problem statement and business objective

Data source and dataset structure

Data cleaning and preprocessing

KPI calculation

Retailer and brand analysis

Regional and state analysis

Monthly and pricing analysis

Visualizations and dashboard

Key insights

Business recommendations and conclusion

A useful storytelling framework is:

Data → Observation → Business Meaning → Recommendation

⚠️ **Limitations**

The dataset covers only 2021, so long-term year-over-year growth
cannot be assessed.

Additional drivers such as advertising spend, promotions,
store-level distribution, competitor pricing, and detailed costs are
not included.

Correlation identifies association and does not establish causation.

Additional years and operational variables would support deeper
forecasting and causal analysis.

🔮 **Future Analysis**

With additional data, the project could be extended to:

Year-over-year sales growth

Sales forecasting

Promotion effectiveness

Store-level analysis

Distribution coverage

Customer segmentation

Price elasticity

Regression analysis

Profitability forecasting

Regional demand forecasting

👤 **Project Context**

This is an Excel-based business analytics project demonstrating the
complete workflow:

Problem Definition → Data Understanding → Data Cleaning → KPI
Development → Exploratory Analysis → Visualization → Insights → Business
Recommendations

⭐ **Key Takeaway**

Coca-Cola should optimize for both sales growth and profitability
rather than focusing on revenue alone.

The strongest sales markets, strongest profit markets, and strongest
margin markets are not always the same. Combining retailer, brand,
geographic, pricing, and seasonal analysis provides a stronger basis for
business decisions.
