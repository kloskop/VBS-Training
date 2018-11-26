
option explicit

Dim total

total = 60000/12
total = total - (total*0.2 + total*0.05 + 200)

MsgBox "Adam recieves $" & total &" as monthly salary after all deductions!", 64, "Salary info"

