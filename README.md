# Profitability-Analysis-using-Monte-Carlo Simulation

## Introduction

A Monte Carlo simulation is a powerful tool for investigating uncertainty in a complex process. It simulates the possible pathway of a future endeavor, experiment, or process, given multiple inputs that each have uncertainty. This project aims to implement a Monte Carlo simulation for financial planning processes.

## Problem Statement

The project involves creating a VBA user form with the following functionalities:

1. **User Input**: Users can input bounds for the various distributions as indicated in the problem statement.
2. **Simulation**: When the GO button is clicked, the simulation works (does not crash) for at least 1,000 simulations.
3. **Output**: The percentage of simulations that resulted in positive net present value (NPV) is output in a message box.
4. **Histogram**: The results of the simulation are output successfully in a histogram.

## How to Check Your Simulation

To determine whether your simulation is working properly, you can test it for a specific set of values. The result of this simulation should be that approximately 40% (+/- 5%) of simulations are profitable.

![image](https://github.com/God-ass/Profitability-Analysis-using-Monte-Carlo/assets/92200827/5a1cb828-fdfd-43b0-9f0d-69787dfcaa1b)

And this is what the user form might look like 

![image](https://github.com/God-ass/Profitability-Analysis-using-Monte-Carlo/assets/92200827/3e2ae1ee-bb0b-4ee2-aad5-4318f1d20156)


## Hints

- If you are getting unexpected results from calculations involving variables that are input into text boxes on a VBA user form, you can trick VBA into thinking it is a number by simply multiplying by 1.
- The Beta_Inv function (borrowed from Excel) cannot have negative arguments. Instead, make the arguments positive, perform the calculation using the Beta_Inv function, then negate the output at the end.

