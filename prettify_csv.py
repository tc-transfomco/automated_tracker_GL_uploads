IN_FILE = "result.csv"
import csv


def main():
    out_file = IN_FILE[:-4] + "_pretty" + IN_FILE[-4:]
    max_widths = []
    with open(IN_FILE, 'r') as f:
        reader = csv.reader(f, delimiter=',', quotechar='\"')
        for row, values in enumerate(reader):
            for i, value in enumerate(values):
                if row == 0:
                    max_widths.append(len(value) + 3)
                    continue
                max_widths[i] = max(len(value) + 3, max_widths[i])

    with open(IN_FILE, 'r') as f:
        reader = csv.reader(f, delimiter=',', quotechar='\"')
        with open(out_file, 'w') as out:
            for values in reader:
                for i, value in enumerate(values):
                    out.write(str(f"{value},").ljust(max_widths[i] - 1))
                out.write('\n')


if __name__ == "__main__":
    main()
