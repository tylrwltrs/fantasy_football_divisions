import pandas
import random
import statistics

input_df = pandas.read_excel("Input.xlsx")

sorted_input = input_df.sort_values(by="PF", ascending=False)

input_list = sorted_input.values.tolist()

target_median = statistics.median(input_df["PF"])
target_avg = input_df.sum()['PF']/12
target_std_dev = statistics.stdev(input_df["PF"])
print("\nTarget Median : ", target_median)
print("Target Average : ", target_avg)
print("Target Standard Deviation : ", target_std_dev, "\n")

best_med_score = .5
good_medians = []
try:
    for i in range(int(input("\n# of random divisions to test?\n"))):
        randomized_indexes = [0,1,2,3,4,5,6,7,8,9,10,11]
        random.shuffle(randomized_indexes)
        
        # use random indexes to assign divisions
        east_division = []
        curr_sum_east = 0
        west_division = []
        curr_sum_west = 0
        for i in randomized_indexes:
            if len(east_division) < 6:
                east_division.append(input_list[i][1])
                curr_sum_east += input_list[i][1]
            else:
                west_division.append(input_list[i][1])
                curr_sum_west += input_list[i][1]
        
        score_east = abs(statistics.median(east_division) - target_median)/target_median
        score_west = abs(statistics.median(west_division) - target_median)/target_median  
        score = score_east + score_west
        if score < best_med_score:
            good_medians.clear()
            good_medians.append(randomized_indexes)
            best_med_score = score
        elif score == best_med_score:
            good_medians.append(randomized_indexes)
    
    best_avg_score = .5
    good_median_and_average = []
    for good_median in good_medians:
        east_division = []
        curr_sum_east = 0
        west_division = []
        curr_sum_west = 0
        for idx in good_median:
            if len(east_division) < 6:
                east_division.append(input_list[idx][1])
                curr_sum_east += input_list[idx][1]
            else:
                west_division.append(input_list[idx][1])
                curr_sum_west += input_list[idx][1]
        score_east = abs(statistics.mean(east_division) - target_avg)/target_avg
        score_west = abs(statistics.mean(west_division) - target_avg)/target_avg
        score = score_east + score_west
        if score <= best_avg_score:
            best_avg_score = score
            good_median_and_average = good_median

except KeyboardInterrupt:
    print("interupted...good enough :)")

print("\nBest Median Score : ", round(best_med_score, ndigits=6), "%"," from the desired target")
print("Best Average Score : ", round(best_avg_score, ndigits=6), "%"," from the desired target\n")
east_division = []
west_division = []
for i in good_median_and_average:
    if len(east_division) < 6:
        east_division.append(input_list[i])
    else:
        west_division.append(input_list[i])
        
east_output = pandas.DataFrame(east_division)
east_output.rename(columns={0:'EAST',1:'PF'}, inplace=True)
east_output.sort_values(by='PF', ascending=False, inplace=True, ignore_index=True)
west_output = pandas.DataFrame(west_division)
west_output.rename(columns={0:'WEST',1:'PF'}, inplace=True)
west_output.sort_values(by='PF', ascending=False, inplace=True, ignore_index=True)
sum_east = sum(east_output['PF'])
sum_west = sum(west_output['PF'])
diff = abs(sum_east - sum_west)

combined = east_output.join(west_output,lsuffix="(E)",rsuffix="(W)")
print(combined.to_string(index=False))

print("\nEAST SUM : ", sum_east)
print("EAST AVG : ", sum_east/6)
print("EAST STD DEV :", statistics.stdev(east_output['PF']))

print("\nWEST SUM : ", sum_west)
print("WEST AVG : ", sum_west/6)
print("WEST STD DEV :", statistics.stdev(west_output['PF']))

print("\nDiff in total sum : ", diff, "\n")

with pandas.ExcelWriter("Output.xlsx") as writer:
    east_output.to_excel(writer, sheet_name="East", index=False, header=False)
    west_output.to_excel(writer, sheet_name="West", index=False, header=False)