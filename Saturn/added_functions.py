


@excel_helper()
def match(lookup_value, lookup_range, match_type=1): # Excel reference: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a

    if list_like(lookup_range) is False:
        return ExcelError('#VALUE!', 'Lookup_range is not a Range')

    def type_convert(value):
        if type(value) == str:
            value = value.lower()
        elif type(value) == int:
            value = float(value)
        elif value is None:
            value = 0

        return value

    def type_convert_float(value):
        if is_number(value):
            value = float(value)
        else:
            value = None

        return value

    lookup_value = type_convert(lookup_value)

    range_values = [item for t in lookup_range for item in t if t is not None]  # filter None values to avoid asc/desc order errors
    # range_values = [x for x in lookup_range[0]) if x is not None] # filter None values to avoid asc/desc order errors
    range_length = len(range_values)

    if match_type == 1:
        # Verify ascending sort

        posMax = -1
        for i in range(range_length):
            current = type_convert(range_values[i])

            if i < range_length - 1:
                if current > type_convert(range_values[i + 1]):
                    return ExcelError('#VALUE!', 'for match_type 1, lookup_range must be sorted ascending')
            if current <= lookup_value:
                posMax = i
        if posMax == -1:
            return ExcelError('#VALUE!','no result in lookup_range for match_type 1')
        return posMax +1 #Excel starts at 1

    elif match_type == 0:
        # No string wildcard
        try:
            if is_number(lookup_value):
                lookup_value = float(lookup_value)
                output = [type_convert_float(x) for x in range_values].index(lookup_value) + 1
            else:
                output = [str(x).lower() for x in range_values].index(lookup_value) + 1
            return output
        except:
            return ExcelError('#VALUE!', '%s not found' % lookup_value)

    elif match_type == -1:
        # Verify descending sort
        posMin = -1
        for i in range((range_length)):
            current = type_convert(range_values[i])

            if i is not range_length-1 and current < type_convert(range_values[i+1]):
               return ExcelError('#VALUE!','for match_type -1, lookup_range must be sorted descending')
            if current >= lookup_value:
               posMin = i
        if posMin == -1:
            return ExcelError('#VALUE!', 'no result in lookup_range for match_type -1')
        return posMin +1 #Excel starts at 1

def choose(index_num, *values): # Excel reference: https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc

    index = int(index_num)

    if index <= 0 or index > 254:
        return ExcelError('#VALUE!', '%s must be between 1 and 254' % str(index_num))
    elif index > len(values):
        return ExcelError('#VALUE!', '%s must not be larger than the number of values: %s' % (str(index_num), len(values)))
    else:
        return values[index - 1]
