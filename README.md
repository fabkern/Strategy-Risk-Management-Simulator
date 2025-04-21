# **Strategy Risk Simulator – Trade Viability Engine**

## Summary  
This project is a forward-simulation engine designed to test the viability of trading strategies based on known win rate, risk-to-reward ratio, capital allocation, and concurrent trade exposure.

Rather than testing strategies on historical data, this simulator uses probabilistic modeling (Monte Carlo simulation) to evaluate how strategies survive across thousands of randomized trade paths. It highlights structural fragility (e.g., drawdown risk, capital erosion, loss streak exposure) and quantifies expected performance ranges—helping validate whether a strategy is *survivable*, not just profitable.

---

## What It Does

| Capability | Description |
|------------|-------------|
| **Monte Carlo Simulation** | Simulates thousands of randomized trade outcomes based on user-defined parameters (win rate, R:R, trade count, concurrent positions)  
| **Concurrency Support** | Models simultaneous trade exposure and adjusts risk calculations accordingly  
| **Fixed or Percent Risk Models** | Supports both fixed-dollar and percentage-based capital exposure modes  
| **Best / Worst / Median Outcomes** | Tracks the highest-performing run, lowest-performing run, and median path across all simulations  
| **Drawdown Analysis** | Calculates max drawdown in both dollar and percentage terms for each path  
| **Streak Analysis** | Logs the longest streaks of consecutive wins and losses  
| **Excel Report Generator** | Outputs simulation results into a color-coded Excel file with embedded performance charts  
| **Visual Output** | Plots the performance of all simulations + overlays best/worst/most likely scenarios

---

## Use Cases

- **Strategy Development**: Test theoretical edge and capital behavior before building a live backtest  
- **Portfolio Stress Testing**: Model edge degradation, volatility clustering, or overexposure risk  
- **Risk Calibration**: Quantify survivability at different risk levels and win probabilities  
- **Visual Risk Communication**: Use Excel output to communicate complex statistical behavior to non-technical stakeholders

