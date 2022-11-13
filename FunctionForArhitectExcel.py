import re
import struct
from math import pow
import global_


def reverse(string):  # returns the reversed string
    string = "".join(reversed(string))
    return string


def hexCorrespond(string):  # this function is used for hex conversion
    if string == 10:
        string = "A"
    elif string == 11:
        string = "B"
    elif string == 12:
        string = "C"
    elif string == 13:
        string = "D"
    elif string == 14:
        string = "E"
    elif string == 15:
        string = "F"
    else:
        return string
    return string


def conversionAndVerifyType(type):  # adapt type parameters from Architect excel to Output excel
    if type.upper() == "UI8" or type.upper() == "UINT8":
        type = "uint8"
        return type

    elif type.upper() == "SI8" or type.upper() == "SINT8":
        type = "sint8"
        return type

    elif type.upper() == "UI16" or type.upper() == "UINT16":
        type = "uint16"
        return type

    elif type.upper() == "SI16" or type.upper() == "SINT16":
        type = "sint16"
        return type

    elif type.upper() == "UI32" or type.upper() == "UINT32":
        type = "uint32"
        return type

    elif type.upper() == "SI32" or type.upper() == "SINT32":
        type = "sint32"
        return type

    elif type.upper() == "FL32" or type.upper()== "FLOAT32" or type.upper() == "float".upper():
        type = "float32"
        return type

    elif type.upper() == "UI64" or type.upper() == "UINT64":
        type = "uint64"
        return type

    elif type.upper() == "SI64" or type.upper() == "SINT64":
        type = "sint64"
        return type

    elif type == "2x SI16":
        type = "sint16"
        return type

    return False


def completeNoneValue(type, value, defaultValue):  # if value is missing, this functions fills the value
    if value is None:
        if "0xFF" in defaultValue:  # DE CE
            if type == "uint8":
                value = "0xFF"
            elif type == "sint8":
                value = "0xFF"
            elif type == "uint16":
                value = "0xFFFF"
            elif type == "sint16":
                value = "0xFFFF"
            elif type == "uint32":
                value = "0xFFFFFFFF"
            elif type == "sint32":
                value = "0xFFFFFFFF"
            elif type == "float32":
                value = "0xFFFFFFFF"
            elif type == "uint64":
                value = "0xFFFFFFFFFFFFFFFF"
            elif type == "sint64":
                value = "0xFFFFFFFFFFFFFFFF"
        else:
            value = defaultValue

    if value == "0x0" or value == "PROD":
        if type == "uint8":
            value = "0x00"
        elif type == "sint8":
            value = "0x00"
        elif type == "uint16":
            value = "0x0000"
        elif type == "sint16":
            value = "0x0000"
        elif type == "uint32":
            value = "0x00000000"
        elif type == "sint32":
            value = "0x00000000"
        elif type == "float32":
            value = "0x00000000"
        elif type == "uint64":
            value = "0x0000000000000000"
        elif type == "sint64":
            value = "0x0000000000000000"

    if value == "NA" or value == "N/A\\0":
        if type == "uint8":
            value = "0xFF"
        elif type == "sint8":
            value = "0xFF"
        elif type == "uint16":
            value = "0xFFFF"
        elif type == "sint16":
            value = "0xFFFF"
        elif type == "uint32":
            value = "0xFFFFFFFF"
        elif type == "sint32":
            value = "0xFFFFFFFF"
        elif type == "float32":
            value = "0xFFFFFFFF"
        elif type == "uint64":
            value = "0xFFFFFFFFFFFFFFFF"
        elif type == "sint64":
            value = "0xFFFFFFFFFFFFFFFF"
    return value


def conversionValue(type, count, value):  # adapt the value from Architect excel for Output excel
    if value is not None:
        if isinstance(value, str):
            value = value.strip()
            if "0X" in value:
                value = value.replace("0X", "0x")
            if (type == "uint8" or type == "sint8") and value[0:2] == "0x" and len(value) < 50 and "," not in value and ";" not in value:
                value = conversionInt8(value)
                if len(value[2:]) < count * 2:
                    value = multiplyHex(value, int(count))
            elif (type == "uint16" or type == "sint16") and value[0:2] == "0x" and len(value) < 50 and "," not in value and ";" not in value:
                value = conversionInt16(value)
                if len(value[2:]) < count * 4:
                    value = multiplyHex(value, int(count))
            elif (type == "uint32" or type == "sint32") and value[0:2] == "0x" and len(value) < 50 and "," not in value and ";" not in value:
                value = conversionInt32(value)
                if len(value[2:]) < count * 8:
                    value = multiplyHex(value, int(count))
            elif "\"" in value:
                value = conversionAsciiToHexSecondMethod(count, value)
            elif not isinstance(value, int) and "0x" not in value and ";" not in value:
                value = conversionAsciiToHexThirdMethod(count, value)
        return value


def conversionInt8(value):  # converts a value to write hex value
    if len(value) <= 4 or int(value, 16) == 0:
        value = int(value, 16)
        string = ""
        while value != 0:
            string = string + str(value % 2)
            value = value // 2
        string = reverse(string)
        string = string.zfill(8)
        string1 = string[0:4]
        string2 = string[4:8]
        string1 = int(string1, 2)
        string2 = int(string2, 2)
        string1 = str(hexCorrespond(string1))
        string2 = str(hexCorrespond(string2))
        value = string1 + string2
        value = "0x" + value
    return value


def multiplyHex(value, count):  # multiplies the value according to the count
    initialValue = value[2:]
    if count != 1:
        for index in range(count - 1):
            value += initialValue
    return value


def conversionInt16(value):  # converts a value to write hex
    if len(value) <= 6 or int(value, 16) == 0:
        value = int(value, 16)
        string = ""
        while value != 0:
            string = string + str(value % 2)
            value = value // 2
        string = reverse(string)
        string = string.zfill(16)
        string1 = string[0:4]
        string2 = string[4:8]
        string3 = string[8:12]
        string4 = string[12:16]
        string1 = int(string1, 2)
        string2 = int(string2, 2)
        string3 = int(string3, 2)
        string4 = int(string4, 2)
        string1 = str(hexCorrespond(string1))
        string2 = str(hexCorrespond(string2))
        string3 = str(hexCorrespond(string3))
        string4 = str(hexCorrespond(string4))
        value = string1 + string2 + string3 + string4
        value = "0x" + value
    return value


def conversionInt32(value):  # converts a value to write hex
    if len(value) <= 10 or int(value, 16) == 0:
        value = int(value, 16)
        string = ""
        while value != 0:
            string = string + str(value % 2)
            value = value // 2
        string = reverse(string)
        string = string.zfill(32)
        string1 = string[0:4]
        string2 = string[4:8]
        string3 = string[8:12]
        string4 = string[12:16]
        string5 = string[16:20]
        string6 = string[20:24]
        string7 = string[24:28]
        string8 = string[28:32]
        string1 = int(string1, 2)
        string2 = int(string2, 2)
        string3 = int(string3, 2)
        string4 = int(string4, 2)
        string5 = int(string5, 2)
        string6 = int(string6, 2)
        string7 = int(string7, 2)
        string8 = int(string8, 2)
        string1 = str(hexCorrespond(string1))
        string2 = str(hexCorrespond(string2))
        string3 = str(hexCorrespond(string3))
        string4 = str(hexCorrespond(string4))
        string5 = str(hexCorrespond(string5))
        string6 = str(hexCorrespond(string6))
        string7 = str(hexCorrespond(string7))
        string8 = str(hexCorrespond(string8))
        value = string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8
        value = "0x" + value
    return value


def conversionInt64(value):  # converts a value to write hex
    if value is not None:
        value = int(value)
        if isinstance(value, int):
            string = ""
            while value != 0:
                string = string + str(value % 2)
                value = value // 2
            string = reverse(string)
            string = string.zfill(64)
            string1 = string[0:4]
            string2 = string[4:8]
            string3 = string[8:12]
            string4 = string[12:16]
            string5 = string[16:20]
            string6 = string[20:24]
            string7 = string[24:28]
            string8 = string[28:32]
            string9 = string[32:36]
            string10 = string[36:40]
            string11 = string[40:44]
            string12 = string[44:48]
            string13 = string[48:52]
            string14 = string[52:56]
            string15 = string[56:60]
            string16 = string[60:64]

            string1 = int(string1, 2)
            string2 = int(string2, 2)
            string3 = int(string3, 2)
            string4 = int(string4, 2)
            string5 = int(string5, 2)
            string6 = int(string6, 2)
            string7 = int(string7, 2)
            string8 = int(string8, 2)
            string9 = int(string9, 2)
            string10 = int(string10, 2)
            string11 = int(string11, 2)
            string12 = int(string12, 2)
            string13 = int(string13, 2)
            string14 = int(string14, 2)
            string15 = int(string15, 2)
            string16 = int(string16, 2)

            string1 = str(hexCorrespond(string1))
            string2 = str(hexCorrespond(string2))
            string3 = str(hexCorrespond(string3))
            string4 = str(hexCorrespond(string4))
            string5 = str(hexCorrespond(string5))
            string6 = str(hexCorrespond(string6))
            string7 = str(hexCorrespond(string7))
            string8 = str(hexCorrespond(string8))
            string9 = str(hexCorrespond(string9))
            string10 = str(hexCorrespond(string10))
            string11 = str(hexCorrespond(string11))
            string12 = str(hexCorrespond(string12))
            string13 = str(hexCorrespond(string13))
            string14 = str(hexCorrespond(string14))
            string15 = str(hexCorrespond(string15))
            string16 = str(hexCorrespond(string16))
            value = string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15 + string16
            value = "0x" + value
            return value


def conversionInt32Address(value):  # converts the address from int to hex
    string = ""
    while value != 0:
        string = string + str(value % 2)
        value = value // 2
    string = reverse(string)
    string = string.zfill(32)
    string1 = string[0:4]
    string2 = string[4:8]
    string3 = string[8:12]
    string4 = string[12:16]
    string5 = string[16:20]
    string6 = string[20:24]
    string7 = string[24:28]
    string8 = string[28:32]
    string1 = int(string1, 2)
    string2 = int(string2, 2)
    string3 = int(string3, 2)
    string4 = int(string4, 2)
    string5 = int(string5, 2)
    string6 = int(string6, 2)
    string7 = int(string7, 2)
    string8 = int(string8, 2)
    string1 = str(hexCorrespond(string1))
    string2 = str(hexCorrespond(string2))
    string3 = str(hexCorrespond(string3))
    string4 = str(hexCorrespond(string4))
    string5 = str(hexCorrespond(string5))
    string6 = str(hexCorrespond(string6))
    string7 = str(hexCorrespond(string7))
    string8 = str(hexCorrespond(string8))
    value = string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8
    contor = 0
    x = value[0]
    while x == "0":
        contor += 1
        x = value[contor]
    value = value[contor:]
    value = "0x" + value
    return value


def conversionAsciiToHex(count, value):  # converts from ascii to hex
    intermediaryValue = ""
    if value[0] == "\"":
        for char in value[1:]:
            if char == "\\":
                break
            intermediaryValue = intermediaryValue + char
        value = intermediaryValue
        value = value.encode('utf-8')
        value = value.hex().upper()
        value = value.ljust(2 * count, '0')
        value = "0x" + value
    else:
        for char in value[0:]:
            if char == "\\":
                break
            intermediaryValue = intermediaryValue + char
        value = intermediaryValue
        value = value.encode('utf-8')
        value = value.hex().upper()
        value = value.ljust(2 * count, '0')
        value = "0x" + value
    return value


def conversionAsciiToHexSecondMethod(count, value):  # converts from ascii to hex
    intermediaryValue = ""
    for char in value[1:]:
        if char == "\"":
            break
        intermediaryValue = intermediaryValue + char
    value = intermediaryValue
    value = value.encode('utf-8')
    value = value.hex().upper()
    value = value.ljust(2 * count, '0')
    value = "0x" + value
    return value


def conversionAsciiToHexThirdMethod(count, value):  # converts from ascii to hex
    value = value.encode('utf-8')
    value = value.hex().upper()
    value = value.ljust(2 * count, '0')
    value = "0x" + value
    return value


def conversionCountFor2Value(type, count):  # adapts the count according to the type
    if type == "2x SI16":
        count = count * 2
    return count


def conversionOem2(value, count):  # converts values into hex
    if "0x" not in value:
        if "[ASCII]" in value:
            index = value.index("[ASCII]")
            value = value[:index]
        value = value.strip()
        if "\"" in value:
            list = []
            for contor in range(len(value)):
                if value[contor] == "\"":
                    list.append(contor)
            value = value[list[0] + 1:list[len(list) - 1]]
        value = value.encode('utf-8')
        value = value.hex().upper()
        value = value.ljust(2 * count, '0')
        value = "0x" + value
    elif "0x" in value:
        value = value.replace(" ", "")
    return value


def conversionOem(value, count):  # converts values into hex
    if "0x" not in value:
        value = value.strip()
        if "\"" in value:
            for index in range(len(value)):
                if value[index] == "\"":
                    value = value[index + 1:]
                    break
        if "\"" in value:
            for index in range(len(value) - 1, 0, -1):
                if value[index] == "\"":
                    value = value[:index]
                    break
        elif "[ASCII]" in value:
            for index in range(len(value)):
                if value[index] == "[":
                    value = value[:index]
                    break
        value = value.strip()
        value = value.encode('utf-8')
        value = value.hex().upper()
        value = value.ljust(2 * count, '0')
        value = "0x" + value
    return value


def functionForListWithNotUsedElements(listWithNotUsedElements):  # splits the elements into a list
    list = re.split(';|,', listWithNotUsedElements)
    for iterator in range(0, len(list)):
        list[iterator] = list[iterator].strip()
    return list


def twosComplement_hex(hexValue):  # converts the hex value to negative number
    bits = 16  # Number of bits in a hexadecimal number format
    val = int(hexValue, bits)
    if val & (1 << (bits - 1)):
        val -= 1 << bits
    return val


def range_verification(type, value):
    if type == "uint8":
        if 255 >= value >= 0:
            return True

    elif type == "sint8":
        if -128 <= value <= 127:
            return True

    elif type == "uint16":
        if 0 <= value <= 65535:
            return True

    elif type == "sint16":
        if -32768 <= value <= 32767:
            return True

    elif type == "uint32":
        if 0 <= value <= 4294967295:
            return True

    elif type == "sint32":
        if -2147483648 <= value <= 2147483647:
            return True

    elif type == "float32":
        if -3.4028235 * pow(10, 38) <= value <= 3.4028235 * pow(10, 38):
            return True

    elif type == "uint64":
        if 0 <= value <= 18446744073709551615:
            return True

    elif type == "sint64":
        if -9223372036854775808 <= value <= 9223372036854775807:
            return True

    return False


def valid_data(type, count, value):
    if (type == "uint8" or type == "sint8") and count == 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if type == "sint8":
                    if range_verification(type, twosComplement_hex(value)) is False:
                        return False
                elif range_verification(type, int(value, 16)) is False:
                    return False

        else:
            if range_verification(type, value) is False:
                return False

    elif (type == "uint8" or type == "sint8") and count > 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if ";" not in value and "," not in value:
                    startString = 0
                    stopString = 2
                    valueAux = value[2:]
                    while valueAux != "":
                        newValue = valueAux[startString:stopString]
                        valueAux = valueAux[stopString:]
                        if type == "sint8":
                            if range_verification(type, twosComplement_hex(newValue)) is False:
                                return False
                        elif range_verification(type, int(newValue, 16)) is False:
                            return False
                else:
                    listWithValue = re.split(';|,', str(value))
                    for valueAux in listWithValue:
                        if type == "sint8":
                            if range_verification(type, twosComplement_hex(valueAux)) is False:
                                return False
                        elif range_verification(type, int(valueAux, 16)) is False:
                            return False
            else:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if range_verification(type, int(valueAux)) is False:
                        return False
        else:
            if isinstance(value, int):
                if range_verification(type, value) is False:
                    return False

    elif (type == "uint16" or type == "sint16") and count == 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if type == "sint16":
                    if range_verification(type, twosComplement_hex(value)) is False:
                        return False
                elif range_verification(type, int(value, 16)) is False:
                    return False
        else:
            if range_verification(type, value) is False:
                return False

    elif (type == "uint16" or type == "sint16") and count > 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if ";" not in value and "," not in value:
                    startString = 0
                    stopString = 4
                    valueAux = value[2:]
                    while valueAux != "":
                        newValue = valueAux[startString:stopString]
                        valueAux = valueAux[stopString:]
                        if type == "sint16":
                            if range_verification(type, twosComplement_hex(newValue)) is False:
                                return False
                        elif range_verification(type, int(newValue, 16)) is False:
                            return False
                else:
                    listWithValue = re.split(';|,', str(value))
                    for valueAux in listWithValue:
                        if type == "sint16":
                            if range_verification(type, twosComplement_hex(valueAux)) is False:
                                print(valueAux + "->")
                                print(twosComplement_hex(valueAux))
                                return False
                        elif range_verification(type, int(valueAux, 16)) is False:
                            return False
            else:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if range_verification(type, int(valueAux)) is False:
                        return False
        else:
            if isinstance(value, int):
                if range_verification(type, value) is False:
                    return False

    elif (type == "uint32" or type == "sint32") and count == 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if type == "sint32":
                    if range_verification(type, twosComplement_hex(value)) is False:
                        return False
                elif range_verification(type, int(value, 16)) is False:
                    return False
        else:
            if range_verification(type, value) is False:
                return False

    elif (type == "uint32" or type == "sint32") and count > 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if ";" not in value and "," not in value:
                    startString = 0
                    stopString = 8
                    valueAux = value[2:]
                    while valueAux != "":
                        newValue = valueAux[startString:stopString]
                        valueAux = valueAux[stopString:]
                        if type == "sint32":
                            if range_verification(type, twosComplement_hex(newValue)) is False:
                                return False
                        elif range_verification(type, int(newValue, 16)) is False:
                            return False

                else:
                    listWithValue = re.split(';|,', str(value))
                    for valueAux in listWithValue:
                        if type == "sint32":
                            if range_verification(type, twosComplement_hex(valueAux)) is False:
                                return False
                        elif range_verification(type, int(valueAux, 16)) is False:
                            return False
            else:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if range_verification(type, int(valueAux)) is False:
                        return False
        else:
            if isinstance(value, int):
                if range_verification(type, value) is False:
                    return False

    elif (type == "uint64" or type == "sint64") and count == 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if type == "sint64":
                    if range_verification(type, twosComplement_hex(value)) is False:
                        return False
                elif range_verification(type, int(value, 16)) is False:
                    return False
        else:
            if range_verification(type, value) is False:
                return False

    elif (type == "uint64" or type == "sint64") and count > 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if ";" not in value and "," not in value:
                    startString = 0
                    stopString = 16
                    valueAux = value[2:]
                    while valueAux != "":
                        newValue = valueAux[startString:stopString]
                        valueAux = valueAux[stopString:]
                        if range_verification(type, int(newValue, 16)) is False:
                            if type == "sint64":
                                if range_verification(type, twosComplement_hex(newValue)) is False:
                                    return False
                            elif range_verification(type, int(newValue, 16)) is False:
                                return False
                else:
                    listWithValue = re.split(';|,', str(value))
                    for valueAux in listWithValue:
                        if range_verification(type, int(valueAux, 16)) is False:
                            if type == "sint64":
                                if range_verification(type, twosComplement_hex(valueAux)) is False:
                                    return False
                            elif range_verification(type, int(valueAux, 16)) is False:
                                return False
            else:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if range_verification(type, int(valueAux)) is False:
                        return False
        else:
            if isinstance(value, int):
                if range_verification(type, value) is False:
                    return False

    elif type == "float32" and count == 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if len(value[2:]) == count * 8:
                    valueAux = struct.unpack('!f', bytes.fromhex(value[2:]))[0]
                    if range_verification(type, valueAux) is False:
                        return False
        else:
            if range_verification(type, float(value)) is False:
                return False

    elif type == "float32" and count > 1:
        if isinstance(value, str):
            if "0x" in value or "0X" in value:
                if ";" not in value and "," not in value:
                    if len(value[2:]) == count * 8:
                        startString = 2
                        stopString = 10
                        for iterator in range(0, count):
                            valueAux = value[startString:stopString]
                            if range_verification(type, struct.unpack('!f', bytes.fromhex(valueAux))[0]) is False:
                                return False
                            startString += 8
                            stopString += 8
                else:
                    listWithValue = re.split(';|,', str(value))
                    for valueAux in listWithValue:
                        if "0x" in valueAux or "0X" in valueAux:
                            if len(value[2:]) == count * 8:
                                if range_verification(type,
                                                      struct.unpack('!f', bytes.fromhex(valueAux[2:]))[0]) is False:
                                    return False
                        else:
                            if range_verification(type, float(valueAux)) is False:
                                return False
            else:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if range_verification(type, float(valueAux)) is False:
                        return False
        else:
            if isinstance(value, int) or isinstance(value, float):
                if range_verification(type, float(value)) is False:
                    return False
    return True


def lengthValues(name, value, count, type, fileWarnings):

    if "8" in type:
        if count == 1:
            if len(value[2:]) != 2:
                fileWarnings.write(
                    "[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(type) + ", Count: " + str(count) +
                    ", Value: " + str(value) + "\n")

        else:
            if ";" in value:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if len(valueAux[2:]) != 2:
                        fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                            name) + ", Type: " + str(type) + ", Count: " + str(count)
                                           + ", Value: " + str(valueAux)
                                           + " IN: " + str(value) + "\n")
                        break
                numberOfElements = str(value).count(";")
                if count != numberOfElements + 1:
                    fileWarnings.write(
                        "[COUNT AND NUMBER OF VALUES DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                            type) + ", Count: " + str(count)
                        + ", Value: " + str(value) + "\n")
            else:
                if len(value[2:]) != count * 2:
                    fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                        name) + ", Type: " + str(type) + ", Count: " + str(count)+ ", Value: " + str(value) +
                                       "\n")

    elif "16" in type:
        if count == 1:
            if len(value[2:]) != 4:
                fileWarnings.write(
                    "[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                        type) + ", Count: " + str(count)
                    + ", Value: " + str(value) + "\n")

        else:
            if ";" in value:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if len(valueAux[2:]) != 4:
                        fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                            name) + ", Type: " + str(type) + ", Count: " + str(count)
                                           + ", Value: " + str(valueAux)
                                           + " IN: " + str(value) + "\n")
                        break
                numberOfElements = str(value).count(";")
                if count != numberOfElements + 1:
                    fileWarnings.write(
                        "[COUNT AND NUMBER OF VALUES DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                            type) + ", Count: " + str(count)
                        + ", Value: " + str(value) + "\n")
            else:
                if len(value[2:]) != count * 4:
                    fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                        name) + ", Type: " + str(type) + ", Count: " + str(count) + ", Value: " + str(value) +
                                       "\n")

    elif "32" in type:
        if count == 1:
            if len(value[2:]) != 8:
                fileWarnings.write(
                    "[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                        type) + ", Count: " + str(count) +
                    ", Value: " + str(value) + "\n")

        else:
            if ";" in value:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if len(valueAux[2:]) != 8:
                        fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                            name) + ", Type: " + str(type) + ", Count: " + str(count)
                                           + ", Value: " + str(valueAux)
                                           + " IN: " + str(value) + "\n")
                        break
                numberOfElements = str(value).count(";")
                if count != numberOfElements + 1:
                    fileWarnings.write(
                        "[COUNT AND NUMBER OF VALUES DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                            type) + ", Count: " + str(count)
                        + ", Value: " + str(value) + "\n")
            else:
                if len(value[2:]) != count * 8:
                    fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                        name) + ", Type: " + str(type) + ", Count: " + str(count) + ", Value: " + str(value) +
                                       "\n")

    elif "64" in type:
        if count == 1:
            if len(value[2:]) != 16:
                fileWarnings.write(
                    "[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                        type) + ", Count: " + str(count) +
                    ", Value: " + str(value) + "\n")

        else:
            if ";" in value:
                listWithValue = re.split(';|,', str(value))
                for valueAux in listWithValue:
                    if len(valueAux[2:]) != 16:
                        fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                            name) + ", Type: " + str(type) + ", Count: " + str(count)
                                           + ", Value: " + str(valueAux)
                                           + " IN: " + str(value) + "\n")
                        break
                numberOfElements = str(value).count(";")
                if count != numberOfElements + 1:
                    fileWarnings.write(
                        "[COUNT AND NUMBER OF VALUES DO NOT CORRESPOND]  Name: " + str(name) + ", Type: " + str(
                            type) + ", Count: " + str(count)
                        + ", Value: " + str(value) + "\n")
            else:
                if len(value[2:]) != count * 16:
                    fileWarnings.write("[TYPE AND VALUE DO NOT CORRESPOND]  Name: " + str(
                        name) + ", Type: " + str(type) + ", Count: " + str(count) + ", Value: " + str(value) +
                                       "\n")

    else:
        print("Type not found")
