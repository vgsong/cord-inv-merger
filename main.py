import time
import PyPDF2
import glob
import pandas as pd
import os
import re
import datetime 


class InvMergerBot:
    def __init__(self, y_per, m_per, wo_records_bol=True):
        self.ROLLFOR_PER = (
                            "DONE", "EXPENSES",
                            "INV", "TS",
                         )

        self.main_dir = f"./invmerger/data/invmerger/{str(y_per)}{str(m_per).zfill(2)}"
        self.wo_records_bol = wo_records_bol 
        self.check_per_dir()


    def check_per_dir(self):
        if not os.path.isdir(self.main_dir):
            print(f"PER not detected in {self.main_dir}\nCreating new per dir")
            os.mkdir(self.main_dir)
            time.sleep(0.5)

            for subdir in self.ROLLFOR_PER:
                os.mkdir(f"{self.main_dir}/{subdir}")

            time.sleep(0.5)
            print(
                    "Month-end DIR created. " \
                    "Make sure you populate TS dir\n" \
                    "with the proper FTE group ./TS/LAST, FIRST\n" \
                    "and reload the script.\n"
                )
            
            return

        else:
            print(f'dir month already exists')

            if not os.listdir(f"{self.main_dir}/TS"):
                print(
                    "TS folder is empty, " \
                    "make sure you add the\n" \
                    "TS folders before running."
                )
                return

            else:
                print("TS FTE list detected in dir!")
                self.member_list = [os.path.basename(x) for x in glob.glob(f"{self.main_dir}/TS/*")]
                print(self.member_list)
        return

    def create_wo_history(self):
        data_to_save = pd.read_csv(f"{self.main_dir}/wo.csv",)
        time_stamp = datetime.datetime.now()

        data_to_save["TIMESTAMP"] = time_stamp.strftime('%y%m%d')
        data_to_save.to_csv(f"./invmerger/data/records/wo_{time_stamp.strftime('%Y')}.csv", mode="a", header=None, index=False)

    def start_wo(self):
        # main() will pick up the invoices in INV dir,
        # use the wo.csv as a mapping to direct which PBC TS to attach
        # the function overwrites the appended INVOICES in the INV dir
        # make sure you have a backup on the 1112 invoices before runnning this function

        wo_cols = (
                   "PRONUM", "INVNUM", "EMPNAME",
                   )

        wo = pd.read_csv(f"{self.main_dir}/wo.csv", names=wo_cols, skiprows=1)

        # creates a temp list of invoices into invoice_wo
        # invoice_wo does not contain duplicated invoice number
        # to avoid including the same TS group twice or more
        invoice_wo = wo["INVNUM"].drop_duplicates()

        # REMINDER: this is the function that merges invoices to signed timesheets
        # the loop is based on the invoice_wo
        # which contains unique in num values
        for invnum in invoice_wo:
            time.sleep(1)
            inv_file = PyPDF2.PdfFileReader(
                open(f"{self.main_dir}/INV/{invnum}.pdf",
                     "rb"),  strict=False)

            t = wo.loc[wo["INVNUM"] == int(invnum)]
            print(f"Work order for invoice {invnum}:\n")
            print(t)
            print(f'\n')

            time.sleep(1)
            for tsname in t["EMPNAME"]:
                merger = PyPDF2.PdfFileMerger(strict=False)
                merger.append(inv_file)
                print(f"starting {tsname} merging")

                for fdir in glob.glob(f"{self.main_dir}/TS/{tsname}/*.pdf"):
                    print(f"merging {invnum} with {os.path.basename(fdir)}")
                    ts_file = PyPDF2.PdfFileReader(open(fdir, "rb"))
                    merger.append(ts_file)
                    del ts_file
                    time.sleep(1)

                for fdir in glob.glob(f"{self.main_dir}/EXPENSES/{tsname}/*.pdf"):
                    print(f"merging {invnum} with {os.path.basename(fdir)}")
                    expense_file = PyPDF2.PdfFileReader(open(fdir, "rb"))
                    merger.append(expense_file)
                    del expense_file
                    time.sleep(1)

                merger.write(f"{self.main_dir}/DONE/{invnum}.pdf")
                time.sleep(1)
                print(f"Merged invoice {invnum} saved")
                print(f"\n")
                print(f"\n")
                time.sleep(1)

            del merger
            del inv_file
        
        if self.wo_records_bol:
            self.create_wo_history()
        else:
            print('wo records not saved!')

    def create_workorder(self):

        def prepare_regex_list():
            # gets the reject pattern based on self.member_list in TS dir
            result = str()

            for n in enumerate(self.member_list):
                if n[0] != len(self.member_list)-1:
                    result += f'{n[1]}|'
                else:
                    result += f'{n[1]}'

            return result


        def pdf_to_str(apdf):
            # gets opened pdf obj from passed argument
            # uses PyPDF2 to parse strings
            # and loop pdfdoc into a reader to get str values
            # returns result
            pdfdoc = PyPDF2.PdfFileReader(apdf)
            result = ''

            for i in range(pdfdoc.numPages // 2):
                current_page = pdfdoc.getPage(i)
                result += current_page.extractText()

            return result

        work_order = {
                    "PRONUM": list(),"INVNUM": list(),"EMPNAME": list(),
                    }
                    
        try:
            regex_member_list_pattern = prepare_regex_list()

        except AttributeError:
            print("Empty member_list detected!")
            print("Please populate member folder with teammeber\nLAST, FIRST name format")
            return

        regex_project_pattern = "1112.\d{3}.\d{3}|" \
                                "1112.\d{3}.00R|" \
                                "1112.00E.\d{3}"

        regex_invoice_pattern ="Invoice No:\s*\d{7}"

        print(self.member_list)

        for invoice_dir in glob.glob(f"{self.main_dir}/INV/*.pdf"):

            pdf_invoice = open(invoice_dir, mode="rb")
            invoice_parsed = pdf_to_str(pdf_invoice)

            project_number_found = re.search(regex_project_pattern, invoice_parsed).group(0)
            invoice_number_found = re.search(regex_invoice_pattern, invoice_parsed).group(0)[-7:]

            if re.search(regex_member_list_pattern, invoice_parsed) is None:
                print(f"invoice {os.path.basename(invoice_dir)} not in Brenton\'s Group")
            else:
                temp_dup_checker = []
                for empname in re.finditer(regex_member_list_pattern, invoice_parsed):
                    # temp_dup_checker.append(empname.group(0)) if empname.group(0) not in temp_dup_checker else False
                    if empname.group(0) not in temp_dup_checker:
                        temp_dup_checker.append(empname.group(0))

                for name in temp_dup_checker:
                    work_order["EMPNAME"].append(name)
                    work_order["PRONUM"].append(project_number_found)
                    work_order["INVNUM"].append(invoice_number_found)

        df = pd.DataFrame.from_dict(work_order)
        print("Workorder output:\n")
        print(df)
        df.to_csv(f"{self.main_dir}/wo.csv", index=False)
        time.sleep(0.5)
        print(f"Workorder created at: {self.main_dir}/wo.csv")


def main():
    INVMERGER = InvMergerBot(2022, 9)
    INVMERGER.create_workorder()
    time.sleep(1)
    # INVMERGER.start_wo()
    

if __name__ == '__main__':
    main()
