import pandas
import random
import statistics

input_df = pandas.read_excel("Input.xlsx")

sorted_input = input_df.sort_values(by="PF", ascending=False)

input_list = sorted_input.values.tolist()

target_median = statistics.median(input_df["PF"])
target_avg = input_df.sum()['PF']/12
target_std_dev = statistics.stdev(input_df["PF"])
print("Target Median : ", target_median)
print("\nTarget Average : ", target_avg)
print("Target Standard Deviation : ", target_std_dev, "\n")

prompt = "Choose desired target:\n" + "[A] = Average\n" + "[S] = Standard Deviation\n" + "[M] = Median\n" + "[Z] = ALL\n" + "[SS] = Sort + Split\n"
target_type = input(prompt)
while target_type.casefold() not in {'a','s','m','b','z','ss'}:
    print("INPUT ERROR")
    target_type = input(prompt)

best_score = 400 # 40 should be the minimum score
if target_type.casefold() == 'ss':
    east_division = []
    west_division = []
    for idx, i in enumerate(input_list):
        if idx%2 == 0:
            east_division.append(i)
        else:
            west_division.append(i)
else:
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
            
            # score = how far current averages are from target average
            match target_type.capitalize():
                case 'A':
                    score_east = abs(statistics.mean(east_division) - target_avg)/target_avg
                    score_west = abs(statistics.mean(west_division) - target_avg)/target_avg
                case 'S':
                    score_east = abs(statistics.stdev(east_division) - target_std_dev)/target_std_dev
                    score_west = abs(statistics.stdev(west_division) - target_std_dev)/target_std_dev
                case 'M':
                    score_east = abs(statistics.median(east_division) - target_median)/target_median
                    score_west = abs(statistics.median(west_division) - target_median)/target_median
                case 'Z':
                    score_east = abs(statistics.stdev(east_division) - target_std_dev)/target_std_dev + abs(statistics.mean(east_division) - target_avg)/target_avg + abs(statistics.median(east_division) - target_median)/target_median
                    score_west = abs(statistics.stdev(west_division) - target_std_dev)/target_std_dev + abs(statistics.mean(west_division) - target_avg)/target_avg + abs(statistics.median(west_division) - target_median)/target_median
            score = score_east + score_west

            if score <= best_score:
                best_indexes = randomized_indexes
                best_score = score
    except KeyboardInterrupt:
        print("interupted...good enough :)")

    print("\nBest Score : ", round(best_score, ndigits=6), "%"," from the desired target\n")
    east_division = []
    west_division = []
    for i in best_indexes:
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