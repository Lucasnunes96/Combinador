def move_item_up(list_box):
    currentRow = list_box.currentRow()
    if currentRow > 0:
        currentItem = list_box.takeItem(currentRow)
        list_box.insertItem(currentRow - 1, currentItem)
        list_box.setCurrentRow(currentRow - 1)


def move_item_down(list_box):
    currentRow = list_box.currentRow()
    if currentRow < list_box.count() - 1:
        currentItem = list_box.takeItem(currentRow)
        list_box.insertItem(currentRow + 1, currentItem)
        list_box.setCurrentRow(currentRow + 1)

def mes_do_ano(data):
    meses = {
         "01": "JANEIRO",
         "02": "FEVEREIRO",
         "03": "MARÇO",
         "04": "ABRIL",
         "05": "MAIO",
         "06": "JUNHO",
         "07": "JULHO",
         "08": "AGOSTO",
         "09": "SETEMBRO",
         "10": "OUTUBRO",
         "11": "NOVEMBRO",
         "12": "DEZEMBRO"
     }
    data = data.split("/")
    return f"{data[0]} DE {meses[data[1]]} DE {data[2]}"