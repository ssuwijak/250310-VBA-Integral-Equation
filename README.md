# 250310 VBA with a Physic Integral Equation
---

use the VBA to calculate the integral equation.

## Steps to Solve
1. simplify the integral term.
2. solve the simplified integral term by a numerical method (Simpson's rule)
3. find the root of the equation *( any lambda value that makes f(lambda) = 0 )* by a numerical method (Bi-section)
4. run and plot the graph.

**note :**
- the lambda cannot be zero (error : divided by zero) 
- the lambda is the wave length, so it should be greater than zero physically.


## Running the Demo
*the demo is on the Youtube: [https://www.youtube.com/watch?v=RasU9lRSQqw](https://www.youtube.com/watch?v=RasU9lRSQqw)*

1. enter the suitable range of lambda to execute the iteration (numerical method)
2. maybe some lambda makes the sinh, cosh and tanh to be infinity, so the equation cannnot be calculated.
3. if we donot know the suitable lambda range, so try to plot the graph to be guildance.
4. refer to the demo, you see the range should be grater than 8.92 ==> 9 m, so try 9 to 500.
5. the bi-section method gets the root is about 183.xxx.
6. done !!!