from fpdf import FPDF, HTMLMixin, __version__ as ver   # установить библиотеку fpdf2
from constants import time_now

class HTML_PDF(FPDF, HTMLMixin):
    pass


# печатаем HTML
def pdf_file():
	with open(f'C:\\Users\\User\\Desktop\\Отчет {time_now}.txt', 'r') as file_txt:
                file = file_txt.read().replace('ID', '<p><p>ID')
	pdf = HTML_PDF() # создаем экземпляр
	# добавляем TTF-шрифты, поддерживающие кириллицу.
	# шрифт PoiretOne
	pdf.add_font("Serif", style="", fname=f"PoiretOne.ttf", uni=True)
	pdf.add_font("Serif", style="B", fname=f"PoiretOne.ttf", uni=True)
	pdf.add_font("Serif", style="I", fname=f"PoiretOne.ttf", uni=True)
	pdf.add_font("Serif", style="BI", fname=f"PoiretOne.ttf", uni=True)
	pdf.set_font("Serif", size=15) # устанавливаем шрифт по умолчанию
	pdf.add_page() # добавляем страницу

	pdf.write_html(file, ul_bullet_char='-', table_line_separators=True)
	pdf.output(f"C:\\Users\\User\\Desktop\\Отчет {time_now} в PDF формате.pdf")






