import xlwings as xw
from itertools import permutations


def calculate_permutations(sentence):
    # Split the sentence into words
    words = sentence.split()

    # Generate permutations of all lengths from 1 to the desired length
    perms = []
    for length in range(1, 3 + 1):
        for p in permutations(words, length):
            perm = " ".join(p)
            if perm not in perms:
                perms.append(perm)

    return perms


def main():
    resultStart = 3
    # Open the Excel file and select the sheet
    filename = r'Search term analysis FR(BROAD).xlsx '
    wb = xw.Book(filename)
    sheet = wb.sheets['ASIN']

    # Extract the search terms from the sheet
    search_terms = sheet.range('F3:F129').value

    # Generate permutations for each search term and write to file
    with open("Keywords coverage.txt", "a",
              encoding="utf-8") as f:
        for term in search_terms:
            perms = calculate_permutations(term)
            for perm in perms:
                sheet['G{0}'.format(resultStart)].value = perm
                resultStart = resultStart + 1


if __name__ == '__main__':
    main()
