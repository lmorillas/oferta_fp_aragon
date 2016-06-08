import xlrd
from xlutils import display


def cell_background_color ( workbook, cell ):
    # Returns TRUE if the given cell from the given workbook has a solid cyan RGB (0,255,255) background.
    # Note that the workbook must be opened with formatting_info = True, i.e.
    #     xlrd.open_workbook(xls_filename, formatting_info=True)
    assert type (cell) is xlrd.sheet.Cell
    assert type (workbook) is xlrd.book.Book

    xf_index = cell.xf_index
    xf_style = workbook.xf_list[xf_index]
    xf_background = xf_style.background

    fill_pattern = xf_background.fill_pattern
    pattern_colour_index = xf_background.pattern_colour_index
    background_colour_index = xf_background.background_colour_index

    pattern_colour = workbook.colour_map[pattern_colour_index]
    background_colour = workbook.colour_map[background_colour_index]

    # If the cell has a solid cyan background, then:
    #  - fill_pattern will be 0x01
    #  - pattern_colour will be cyan (0,255,255)
    #  - background_colour is not used with fill pattern 0x01. (undefined value)
    #    So despite the name, for a solid fill, the background colour is not actually the background colour.
    # Refer https://www.openoffice.org/sc/excelfileformat.pdf S. 2.5.12 'Patterns for Cell and Chart Background Area'
    if fill_pattern == 0x01 and pattern_colour == (0,255,255):
        return True
    return False

def esblanco ( workbook, cell ):
    # Returns TRUE if the given cell from the given workbook has a solid cyan RGB (0,255,255) background.
    # Note that the workbook must be opened with formatting_info = True, i.e.
    #     xlrd.open_workbook(xls_filename, formatting_info=True)
    assert type (cell) is xlrd.sheet.Cell
    assert type (workbook) is xlrd.book.Book

    xf_index = cell.xf_index
    xf_style = workbook.xf_list[xf_index]
    xf_background = xf_style.background

    fill_pattern = xf_background.fill_pattern
    pattern_colour_index = xf_background.pattern_colour_index
    background_colour_index = xf_background.background_colour_index

    pattern_colour = workbook.colour_map[pattern_colour_index]
    background_colour = workbook.colour_map[background_colour_index]

    # If the cell has a solid cyan background, then:
    #  - fill_pattern will be 0x01
    #  - pattern_colour will be cyan (0,255,255)
    #  - background_colour is not used with fill pattern 0x01. (undefined value)
    #    So despite the name, for a solid fill, the background colour is not actually the background colour.
    # Refer https://www.openoffice.org/sc/excelfileformat.pdf S. 2.5.12 'Patterns for Cell and Chart Background Area'
    # print fill_pattern, pattern_colour, background_colour
    return not background_colour and not pattern_colour



def limpia_nombre(nombre):
    '''
    IES Hermanos Argensola
BARBASTRO
'''
    nombre = nombre.replace('\n', ' ')
    nombre = nombre.split(' ')
    nombre = [n for n in nombre if not n.isupper()]
    return ' '.join(nombre).strip()

def get_centros(sh, row=0, coli=4):
    while True:
        try:
            cen = sh.cell_value(row, coli)
            if cen:
                yield limpia_nombre(cen)
                coli += 1
            else:
                break
        except IndexError:
            break

def familia_prof(sh, rowi=2, col=1):
    familia=''
    while True:
        if not sh.cell_value(rowi, col+1):
            break
        familia_ = sh.cell_value(rowi, col)
        if familia_:
            familia = familia_
        print familia
        yield {'familia': familia, 'codigo': sh.cell_value(rowi, col+1), 'label': sh.cell_value(rowi, col+2)}
        rowi += 1

def ciclo_for(sh, rowi=2, col=2):
    lista_centros = [c for c in get_centros(sh)]
    while True:
        ciclo = sh.cell_value(rowi, col), sh.cell_value(rowi, col+1)
        if not ciclo[0]:
            return
        #print ciclo
        centros = [c for c in centros_ciclo(sh, lista_centros, rowi)]
        yield({'codigo': ciclo[0],
        'nombre': ciclo[1],
        'centros': centros })
        rowi += 1

def cuenta_centros(sh, row=0, coli=4):
    ini = coli
    while True:
        cen = sh.cell_value(row, coli)
        if cen:
            coli += 1
        else:
            return coli - ini

def centros_ciclo(sh, centros, row):
    coli = 4
    for i in range(len(centros)):
        cell = sh.cell(row, coli + i)
        blanco = esblanco(sh.book, cell)
        if not blanco:
            yield centros[i]

if __name__ == '__main__':
    datos = []
    book = xlrd.open_workbook('GradoSuperior.xls', formatting_info=True)

    '''
    for x in range(0, 13, 2):
        sh = book.sheet_by_index(x)
        #lista_centros = [c for c in get_centros(sh)]
        datos.extend([c for c in ciclo_for(sh)])

    ddatos = {}
    for d in datos:
        codigo = d.get('codigo')
        if ddatos.get(codigo):
            ddatos[codigo]['centros'].extend(d.get('centros'))
        else:
            ddatos[codigo] = {'centros': d.get('centros')}
    '''
    dciclos = []
    for x in range(0, 13, 2):
        sh = book.sheet_by_index(x)
        #lista_centros = [c for c in get_centros(sh)]
        dciclos.extend([c for c in familia_prof(sh)])


