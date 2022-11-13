import logging
import global_

logging.basicConfig(filename='LogError.log', filemode='a', format='%(asctime)s - %(message)s',
                    datefmt='%d-%b-%y %H:%M:%S')  # The format of every line in the log.


def logArchitect():  # pragma: no cover
    logging.warning("ARCHITECT PARAMETERS:")
    logging.warning("Name column number: " + str(global_.nameColumn))
    logging.warning("Type column number: " + str(global_.typeColumn))
    logging.warning("Count column number: " + str(global_.countColumn))
    logging.warning("Value column number: " + str(global_.valueColumn))
    logging.warning("Row start number: " + str(global_.row))
    logging.warning("Row end: " + str(global_.rowEnd))
    logging.warning("Row address number: " + str(global_.readRowAddress))
    logging.warning("Column address number: " + str(global_.readColumnAddress))
    logging.warning("Reserved column number: " + str(global_.resColumn))
    logging.warning("Row a2l number: " + str(global_.rowA2L))
    logging.warning("Column a2l number: " + str(global_.columnA2L) + "\n")


def logOem():    # pragma: no cover
    if global_.checkOem:
        logging.warning("OEM PARAMETERS:")
        logging.warning("Project oem name: " + str(global_.projectOemName))
        logging.warning("Row start number: " + str(global_.rowOem))
        logging.warning("Row end: " + str(global_.rowEndOem))
        logging.warning("Name column number: " + str(global_.nameColumnOem))
        logging.warning("Count column number: " + str(global_.countColumnOem))
        logging.warning("Value column number: " + str(global_.valueColumn))
        logging.warning("Row a2l column: " + str(global_.rowOEMa2l))
        logging.warning("Column a2l column: " + str(global_.columnOemA2L) + "\n")


def logSystem():     # pragma: no cover
    logging.warning("SYSTEM PARAMETERS:")
    logging.warning("Name column number: " + str(global_.nameColumn2))
    logging.warning("Type column number: " + str(global_.typeColumn2))
    logging.warning("Count column number: " + str(global_.countColumn2))
    logging.warning("Value column number: " + str(global_.valueColumn2))
    logging.warning("Row start number: " + str(global_.row2))
    logging.warning("Row end number: " + str(global_.rowEnd2))
    logging.warning("Project specific name: " + str(global_.projectSpecificName))
    logging.warning("Project name type: " + str(global_.projectNameType))
    logging.warning("Elements to be ignored: " + str(global_.listWithNotUsedElements) + "\n")


def logOutput():     # pragma: no cover
    logging.warning("OUTPUT PARAMETERS:")
    logging.warning("Name column number: " + str(global_.nameColumnWrite))
    logging.warning("Type column number: " + str(global_.typeColumnWrite))
    logging.warning("Count column number: " + str(global_.countColumnWrite))
    logging.warning("Value column number: " + str(global_.valueColumnWrite))
    logging.warning("Row start number: " + str(global_.rowWrite))
    logging.warning("L2 write column: " + str(global_.writeL2Architect))
    logging.warning("Tree level 2 write column: " + str(global_.writeTreeLevel2))
    logging.warning("Reserved column: " + str(global_.columnRes))
    logging.warning("Row write address: " + str(global_.writeRowAddress))
    logging.warning("Write column write address: " + str(global_.writeColumnAddress))
    logging.warning("Row write column: " + str(global_.rowWriteA2L))
    logging.warning("Column write column: " + str(global_.columnWriteA2L))
    logging.warning("Column write second sheet column: " + str(global_.columnWriteSheet2))
    logging.warning("Limit value column number: " + str(global_.limitValueColumn))
    logging.warning("Row write second sheet column: " + str(global_.rowWriteSheet2))
    logging.warning("Low limit column number: " + str(global_.lowLimitColumn))
    logging.warning("Max limit column number: " + str(global_.maxLimitColumn) + "\n")
