import os
import textwrap
import sys

import urllib
import urllib.request as ur

from copy import deepcopy

import pandas
import requests
from lxml import etree

from html2image import Html2Image

from PIL import Image, ImageFile

cards_per_sheet = 50
column_count = 10

def svg_tag(basic_tag):
    return '{http://www.w3.org/2000/svg}' + basic_tag


def line_wrap_text(title_element, text_element, text):
    x = text_element.attrib['x']
    max_width = len(title_element.tail)
    title_element.tail = ''
    lines = textwrap.wrap(text=text, width=max_width)

    for i, line in enumerate(lines):
        tspan = etree.SubElement(text_element, "tspan")
        tspan.text = line
        tspan.attrib['x'] = x
        if i > 0:
            tspan.attrib['dy'] = '1em'


def get_image_size(uri):
    # get file size *and* image size (None if not known)
    file = requests.get(uri, stream=True)
    p = ImageFile.Parser()
    for chunk in file.iter_content(chunk_size=1024):
        if not chunk:
            break
        p.feed(chunk)
        if p.image:
            return p.image.size
    file.close()
    return None


def modify_card(card_element, data, card_num):
    parent_map = {c: p for p in card_element.iter() for c in p}

    for title_element in card_element.iterfind(f'.//{svg_tag("title")}'):
        key: str = title_element.text
        if key in data:
            val = data[key]
            parent_element = parent_map[title_element]
            if pandas.isna(val):
                parent_map[parent_element].remove(parent_element)
                continue
            if parent_element.tag == svg_tag('text'):
                parent_element.text = ''
                if 'id' in parent_element.attrib and parent_element.attrib['id'] == 'WRAP':
                    line_wrap_text(title_element, parent_element, val)
                else:
                    title_element.tail = val
            elif parent_element.tag == svg_tag('foreignObject'):
                div_element = parent_element.find("{http://www.w3.org/1999/xhtml}div")
                if div_element is not None:
                    if type(val) == int and div_element.text[0] == '0':
                        val = '0' * max(0, len(div_element.text) - len(str(val))) + str(val)
                    subtree = etree.fromstring(f'<div>{val}</div>')
                    div_element.append(subtree)
                    div_element.text = ""
            elif parent_element.tag == svg_tag('image'):
                parent_element.attrib['{http://www.w3.org/1999/xlink}href'] = f'{str(val).lower()}'
                # scale to fit and preserve aspect ratio
                image_size = get_image_size(str(val).lower())
                width = float(parent_element.attrib['width'])
                height = float(parent_element.attrib['height'])
                scale = max(width / image_size[0], height / image_size[1])

                desired_width = scale * image_size[0]
                desired_height = scale * image_size[1]
                parent_element.attrib['width'] = f'{desired_width}'
                parent_element.attrib['height'] = f'{desired_height}'
            elif parent_element.tag == svg_tag('use'):
                if '{http://www.w3.org/1999/xlink}href' in parent_element.attrib:
                    parent_element.attrib['{http://www.w3.org/1999/xlink}href'] = f'#{str(val).lower()}'

        elif key == 'CardNum':
            parent_element = parent_map[title_element]
            parent_element.text = ''
            title_element.tail = str(card_num)


def make_sheet(name, df):
    filename = f'assets/{name}.svg'
    try:
        tree = etree.parse(filename)
    except:
        print(f"Couldn't open {filename}. Skipping.")
        return
    template_root = tree.getroot()
    template_card = template_root.find(svg_tag('g'))

    # scale viewport to fit all the cards
    (view_x, view_y, card_w, card_h) = template_root.attrib['viewBox'].split()

    card_index = 0

    result_root = deepcopy(template_root)
    result_root.remove(result_root.find(svg_tag('g')))

    # create all the cards and position them correctly
    _: int
    row = 0
    for _, data in df.iterrows():
        count = 1
        if 'copies' in data:
            count = data.copies

        for copy in range(0, count):
            result_root.append(deepcopy(template_card))
            card = result_root[-1]

            sheet_index: int = (card_index % cards_per_sheet)
            column: int = sheet_index % column_count
            row = sheet_index // column_count

            card.attrib['transform'] = f'translate({float(card_w) * column}, {float(card_h) * row})'

            modify_card(card, data, card_index + 1)

            card_index += 1

            if sheet_index + 1 == cards_per_sheet:
                # sheet capacity reached. export it.
                export_sheet(
                    view_x=view_x,
                    view_y=view_y,
                    card_w=card_w,
                    card_h=card_h,
                    column_count=column_count,
                    row_count=row + 1,
                    svg=result_root,
                    sheet_name=f'{name}{(card_index - 1) // cards_per_sheet}'
                )
                # reset svg so that we can generate the next sheet
                result_root = deepcopy(template_root)
                result_root.remove(result_root.find(svg_tag('g')))

    # export remaining cards to it's own sheet
    if (card_index % cards_per_sheet) != cards_per_sheet:
        output_file_name = f'{name}{(card_index - 1) // cards_per_sheet}'
        export_sheet(
            view_x=view_x,
            view_y=view_y,
            card_w=card_w,
            card_h=card_h,
            column_count=column_count,
            row_count=row + 1,
            svg=result_root,
            sheet_name=output_file_name
        )


def export_sheet(view_x, view_y, card_w, card_h, column_count, row_count, svg, sheet_name):
    view_w = float(card_w) * column_count
    view_h = float(card_h) * row_count
    svg.attrib['viewBox'] = f'{view_x}, {view_y}, {view_w}, {view_h}'
    filename = f'output/{sheet_name}.svg'
    print(f'Exporting sheet: {filename}')
    etree.ElementTree(svg).write(filename, pretty_print=True)

    hti = Html2Image(size=(1920*2, 1080*2))
    hti.screenshot(other_file=filename, save_as=f'temp.png')
    im = Image.open(f'temp.png')
    im.getbbox()
    im2 = im.crop(im.getbbox())
    im2.save(f'output/{sheet_name}.png')
    os.remove('temp.png')


def main():
    sheets = pandas.read_excel(f'assets/CardList.xlsx', sheet_name=None)
    for name, sheet in sheets.items():
        make_sheet(name=name, df=sheet)


if __name__ == '__main__':
    main()
