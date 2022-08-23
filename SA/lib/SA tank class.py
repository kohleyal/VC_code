import csv
'''class for the sartorious tanks. originaly writen just for grabing the daily sample data. which is why it leans towards only doing
    that. The class was pulled out of the orginal file "SA_get_daily_sample_data".'''
class Sartorious_tank:
    def __init__(self, tank_id, columns_to_grab=[], eft_col=3, eft_row=12, eft_csv_col=2):
        self.tank_id = tank_id                  #tank label and matches excel worksheet name.
        self.columns_to_grab = columns_to_grab  #columns from csv to pull and in order desired.
        self.eft_col = eft_col                  #eft col in offline data
        self.eft_row = eft_row                  #eft row in offline data
        self.eft_csv_col = eft_csv_col          #eft col in csv raw data
        self.runsheet_label = tank_id           #incase tank id and runsheet don't match in future. change worksheet name here.
        self.batch_run_file = ''
        self.offgas_folder = ''

    def get_daily_data (self):
        #open the csv output file and read through it once. grab the header rows and rows that match the sample eft.
        new_file = []
        with open(self.batch_run_file, "r" ) as tank_file : 
            reader = csv.reader(tank_file, delimiter=";")
            for row in reader :
                if row and self.eft_times:  #run if there is something in the row and EFT timpoints left. could do break after last eft timepoint pop.
                    row_process_time = row[self.eft_csv_col]
                    row = [row[i] for i in self.columns_to_grab]
                    if row_process_time.isalpha():
                        new_file.append(row)
                    elif float(row_process_time) >= self.eft_times[0]:
                        new_file.append(row)
                        self.eft_times.pop(0)
        new_file.append([" "])                
        return new_file

    def grab_offline_eft_times(self, work_sheet_offline):
        print(f'Pulling offline eft for {self.tank_id}')
        offline_eft_list = []
        #generator of eft column cells. returns tuple of input columns which is one in this case.
        check_values = work_sheet_offline.iter_rows(min_row=self.eft_row, max_row=None, min_col=self.eft_col, max_col=self.eft_col, values_only=True)
        for possible_eft in check_values:
            if isinstance(possible_eft[0], (int,float)) and possible_eft[0] >= 0: 
                offline_eft_list.append(possible_eft[0])
            else:
                return offline_eft_list

def main():
    pass

if __name__ == "__main__":
    main()
