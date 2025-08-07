import tkinter as tk
from tkinter import ttk, simpledialog, filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_JUSTIFY, TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
import textwrap
from logos import aplicar_icone

# Função utilitária para quebra de texto
def wrap_text(text, max_width, font_name, font_size):
    """
    Quebra `text` em várias linhas de modo que cada linha caiba em max_width.
    """
    lines, current = [], ""
    for word in text.split():
        test = (current + ' ' + word).strip()
        if stringWidth(test, font_name, font_size) <= max_width:
            current = test
        else:
            if stringWidth(word, font_name, font_size) > max_width:
                if current:
                    lines.append(current)
                    current = ''
                parts = textwrap.wrap(word, width=30)
                lines.extend(parts)
            else:
                lines.append(current)
                current = word
    if current:
        lines.append(current)
    return lines

class CustomStyles:
    @staticmethod
    def apply(root):
        style = ttk.Style(root)
        for theme in ('clam', 'alt', 'vista'):
            try:
                style.theme_use(theme)
                break
            except tk.TclError:
                continue
        base_font = ('Segoe UI', 10)
        header_font = ('Segoe UI', 12, 'bold')
        notebook_bg = '#f5f5f5'
        accent = '#1F4E79'
        style.configure('TFrame', background=notebook_bg)
        style.configure('TNotebook', background=notebook_bg, borderwidth=0)
        style.configure('TNotebook.Tab', font=base_font, padding=(10, 5))
        style.map('TNotebook.Tab', background=[('selected', 'white')])
        style.configure('TLabel', background=notebook_bg, font=base_font)
        style.configure('TEntry', fieldbackground='white', font=base_font)
        style.configure('Accent.TButton', font=base_font, foreground='white', background=accent)
        style.map('Accent.TButton', background=[('active', colors.HexColor('#385a7c')), ('!disabled', accent)])
        style.configure('Treeview', font=base_font, rowheight=48, background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=header_font)

class ScrollableFrame(ttk.Frame):
    """Frame com barra de rolagem vertical e suporte ao scroll do mouse."""
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')

        self.scrollable_frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.window, width=e.width))
        self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)

    def _on_mousewheel(self, event):
        delta = -1 * (event.delta // 120) if hasattr(event, 'delta') else 0
        self.canvas.yview_scroll(delta, 'units')

class EntryDialog(simpledialog.Dialog):
    def __init__(self, parent, title, fields):
        self.fields, self.values = fields, {}
        super().__init__(parent, title)

    def body(self, master):
        # pega o Toplevel real do diálogo
        dlg = self.winfo_toplevel()
        caminho_icone = r"C:\Sistema\logos\Curriculo.ico"
        try:
            aplicar_icone(dlg, caminho_icone)
        except Exception as e:
            print("Ícone não encontrado no diálogo:", e)
            
        self.entries = {}
        pad = {'padx': 5, 'pady': 3}
        for i, (label, multi) in enumerate(self.fields.items()):
            ttk.Label(master, text=label).grid(row=i*2, column=0, sticky='w', **pad)
            if multi:
                frame = ttk.Frame(master)
                frame.grid(row=i*2+1, column=0, **pad)
                text = tk.Text(frame, width=40, height=4, wrap='word')
                vsb = ttk.Scrollbar(frame, orient='vertical', command=text.yview)
                text.configure(yscrollcommand=vsb.set)
                text.pack(side='left', fill='both', expand=True)
                vsb.pack(side='right', fill='y')
                w = text
            else:
                w = ttk.Entry(master, width=40)
                w.grid(row=i*2+1, column=0, **pad)
            self.entries[label] = w
        return next(iter(self.entries.values()))

    def apply(self):
        for lbl, w in self.entries.items():
            self.values[lbl] = w.get('1.0', 'end').strip() if isinstance(w, tk.Text) else w.get().strip()

class SectionFrame(ttk.Frame):
    def __init__(self, parent, title, fields):
        super().__init__(parent)
        self.frame_title = title
        self.fields, self.items = fields, []
        ttk.Label(self, text=title, font=('Segoe UI', 11, 'bold')).pack(anchor='w', pady=(5,0))
        # Treeview com scrollbars
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True)
        vsb = ttk.Scrollbar(container, orient='vertical')
        hsb = ttk.Scrollbar(container, orient='horizontal')
        self.tree = ttk.Treeview(container, columns=list(fields.keys()), show='headings', height=4,
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)
        for c in fields.keys():
            self.tree.heading(c, text=c)
            self.tree.column(c, width=150, anchor='w', stretch=True)
        # Botões
        btns = ttk.Frame(self)
        btns.pack(fill='x', pady=5)
        ttk.Button(btns, text='Adicionar', style='Accent.TButton', command=self.add_item).pack(side='left', padx=2)
        ttk.Button(btns, text='Remover', style='Accent.TButton', command=self.remove_item).pack(side='left', padx=2)

    def add_item(self):
        dlg = EntryDialog(self, f"Adicionar em {self.frame_title}", self.fields)
        if dlg.values:
            self.items.append(dlg.values)
            vals = [dlg.values[k] for k in self.fields.keys()]
            self.tree.insert('', 'end', values=vals)

    def remove_item(self):
        sel = self.tree.selection()
        if sel:
            idx = self.tree.index(sel[0])
            self.tree.delete(sel[0])
            self.items.pop(idx)

class CVGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Currículos")
        self.state('zoomed')
        self.geometry('900x650')
        CustomStyles.apply(self)
        caminho_icone = "C:\\Sistema\\logos\\Curriculo.ico"
        try:
            aplicar_icone(self, caminho_icone)
        except Exception as e:
            print("Ícone não encontrado:", e)

        self._create_widgets()

    def _create_widgets(self):
        scroll = ScrollableFrame(self)
        scroll.pack(fill='both', expand=True)
        parent = scroll.scrollable_frame

        container = ttk.Frame(parent, padding=10)
        container.pack(fill='both', expand=True)
        container.columnconfigure((0,1), weight=1)
        container.rowconfigure(0, weight=1)

        notebook = ttk.Notebook(container)
        notebook.grid(row=0, column=0, columnspan=2, sticky='nsew')

        # Aba Dados Pessoais
        personal = ttk.Frame(notebook, padding=10)
        personal.columnconfigure(1, weight=1)
        labels = ['Nome Completo', 'Data de Nascimento', 'Estado Civil', 'Carteira de Motorista',
                  'Cargo / Título', 'Localização', 'E-mail', 'Telefone', 'LinkedIn', 'GitHub']
        self.entries = {}
        for i, lbl in enumerate(labels):
            ttk.Label(personal, text=lbl).grid(row=i, column=0, sticky='w', pady=5)
            ent = ttk.Entry(personal)
            ent.grid(row=i, column=1, sticky='ew', pady=5)
            self.entries[lbl] = ent
        notebook.add(personal, text='Dados Pessoais')

        # Aba Conteúdo para Currículo
        content = ttk.Frame(notebook, padding=10)
        content.columnconfigure((0,1,2), weight=1)
        ttk.Label(content, text='Resumo', font=('Segoe UI', 12, 'bold')).grid(row=0, column=0, sticky='w')
        self.summary = tk.Text(content, height=5, wrap='word')
        self.summary.grid(row=1, column=0, columnspan=3, sticky='ew', pady=5)

        self.exp = SectionFrame(content, 'Experiência Profissional',
                                {'Empresa': False, 'Cargo': False, 'Período': False, 'Descrição': True})
        self.exp.grid(row=2, column=0, columnspan=3, sticky='nsew')

        self.edu = SectionFrame(content, 'Formação Acadêmica',
                                {'Escola': False, 'Curso': False, 'Período': False, 'Descrição': True})
        self.edu.grid(row=3, column=0, columnspan=3, sticky='nsew')

        self.skills = SectionFrame(content, 'Habilidades', {'Habilidade': False})
        self.skills.grid(row=4, column=0, sticky='nsew')
        self.tech   = SectionFrame(content, 'Conhecimento Técnico', {'Conhecimento': False})
        self.tech.grid(row=4, column=1, sticky='nsew')
        self.lang   = SectionFrame(content, 'Idiomas', {'Idioma': False, 'Nível': False})
        self.lang.grid(row=4, column=2, sticky='nsew')

        notebook.add(content, text='Conteúdo')

        # Aba Carta de Apresentação
        cover_tab = ttk.Frame(notebook, padding=10)
        cover_tab.columnconfigure(1, weight=1)
        cover_labels = ['Destinatário', 'Empresa', 'Data', 'Saudação']
        self.cover_entries = {}
        for i, lbl in enumerate(cover_labels):
            ttk.Label(cover_tab, text=lbl).grid(row=i, column=0, sticky='w', pady=5)
            ent = ttk.Entry(cover_tab)
            ent.grid(row=i, column=1, sticky='ew', pady=5)
            self.cover_entries[lbl] = ent

        ttk.Label(cover_tab, text='Corpo da Carta', font=('Segoe UI', 12, 'bold')).grid(
            row=len(cover_labels), column=0, columnspan=2, sticky='w', pady=(10,2))
        self.cover_body = tk.Text(cover_tab, height=10, wrap='word')
        self.cover_body.grid(row=len(cover_labels)+1, column=0, columnspan=2, sticky='ew', pady=5)

        ttk.Label(cover_tab, text='Despedida e Assinatura', font=('Segoe UI', 12, 'bold')).grid(
            row=len(cover_labels)+2, column=0, columnspan=2, sticky='w', pady=(10,2))
        self.cover_closing = tk.Text(cover_tab, height=4, wrap='word')
        self.cover_closing.grid(row=len(cover_labels)+3, column=0, columnspan=2, sticky='ew', pady=5)

        notebook.add(cover_tab, text='Carta de Apresentação')

        # Botões de geração
        btn_cv = ttk.Button(container, text='Gerar Currículo', style='Accent.TButton',
                            command=self._generate_pdf)
        btn_cv.grid(row=1, column=0, pady=10)

        btn_cover = ttk.Button(container, text='Gerar Carta', style='Accent.TButton',
                               command=self._generate_cover_letter)
        btn_cover.grid(row=1, column=1, pady=10)

    def _generate_pdf(self):
        # coleta dados
        data = {k: e.get().strip() for k, e in self.entries.items()}
        data['Resumo Profissional'] = self.summary.get('1.0', 'end').strip()

        path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF files','*.pdf')])
        if not path:
            return

        c = canvas.Canvas(path, pagesize=A4)
        w, h = A4
        margin    = 2*cm
        sidebar_w = 5*cm

        # largura útil para wrap
        max_content_width = 18 * cm
        raw_width = w - sidebar_w - margin*2 - 10  # cálculo original com buffer
        available_width = min(raw_width, max_content_width)

        # ====== SIDEBAR ======
        ICON_SIZE = 0.5 * cm
        ICONS = {
            'Resumo':      'C:\\Sistema\\logos\\Resumo.png',
            'Experiencia': 'C:\\Sistema\\logos\\Experiencia Profissional.png',
            'Formacao':    'C:\\Sistema\\logos\\Formação.png'
        }

        # Fundo azul lateral
        c.setFillColor(colors.HexColor("#1F4E79"))
        c.rect(0, 0, sidebar_w, h, fill=1, stroke=0)

        # Nome completo no topo
        y = h - margin
        c.setFillColor(colors.white)
        name_font = 'Helvetica-Bold'
        name_font_size = 14
        max_name_width = sidebar_w - 1*cm
        name_lines = wrap_text(data['Nome Completo'], max_name_width, name_font, name_font_size)
        c.setFont(name_font, name_font_size)
        for line in name_lines:
            c.drawString(0.5*cm, y, line)
            y -= (name_font_size + 2)

        # Título "Dados Pessoais"
        y -= 10
        c.setFont('Helvetica-Bold', 12)
        c.drawString(0.5*cm, y, 'Dados Pessoais')

        # Informações abaixo do título, em preto
        y -= 16
        c.setFillColor(colors.white)
        c.setFont('Helvetica', 9)

        fields = [
            ('Data de Nascimento',  data['Data de Nascimento']),
            ('Estado Civil',        data['Estado Civil']),
            ('CNH',                 data['Carteira de Motorista']),
            ('Localização',         data['Localização']),
            ('E-mail',              data['E-mail']),
            ('Telefone',            data['Telefone']),
            ('LinkedIn',            data['LinkedIn']),
            ('GitHub',              data['GitHub']),
        ]

        # filtra só os que não estão em branco
        fields = [(lbl, val) for lbl, val in fields if val.strip()]

        max_label_width = sidebar_w - 1*cm
        for label, val in fields:
            prefix = f"{label}: "
            if label in ('LinkedIn', 'GitHub'):
                # Quebra em linhas que cabem
                lines = wrap_text(prefix + val, max_label_width, 'Helvetica', 9)
                for line in lines:
                    # desenha o texto daquele pedaço
                    c.drawString(0.5*cm, y, line)

                    # separa prefixo de URL (na primeira linha) ou só URL (linhas seguintes)
                    if ': ' in line:
                        pre, url_part = line.split(': ', 1)
                        width_pre = stringWidth(pre + ': ', 'Helvetica', 9)
                    else:
                        url_part = line
                        width_pre = 0

                    # calcula onde começa/finaliza o link naquele pedaço
                    x_start = 0.5*cm + width_pre
                    w_url   = stringWidth(url_part, 'Helvetica', 9)
                    x_end   = x_start + w_url
                    y1, y2  = y - 1, y + 9

                    # cria o link para a URL completa
                    c.linkURL(val, (x_start, y1, x_end, y2), relative=0)

                    y -= 12  # pula para próxima linha
                continue

            # campos normais (sem link)
            lines = wrap_text(prefix + val, max_label_width, 'Helvetica', 9)
            for line in lines:
                c.drawString(0.5*cm, y, line)
                y -= 12

        # Habilidades
        y -= 12
        c.setFont('Helvetica-Bold', 11)
        c.drawString(0.5*cm, y, 'Habilidades')
        c.setFont('Helvetica', 9)
        for skill in self.skills.items:
            skill_line = f'• {skill["Habilidade"]}'
            wrapped_skill = wrap_text(skill_line, max_label_width, 'Helvetica', 9)
            for part in wrapped_skill:
                y -= 12
                c.drawString(0.5*cm, y, part)

        # Conhecimento Técnico
        y -= 18
        c.setFont('Helvetica-Bold', 11)
        c.drawString(0.5*cm, y, 'Conhecimento Técnico')
        c.setFont('Helvetica', 9)
        for tech in self.tech.items:
            tech_line = f'• {tech["Conhecimento"]}'
            wrapped_tech = wrap_text(tech_line, max_label_width, 'Helvetica', 9)
            for part in wrapped_tech:
                y -= 12
                c.drawString(0.5*cm, y, part)

        # Idiomas
        y -= 18
        c.setFont('Helvetica-Bold', 11)
        c.drawString(0.5*cm, y, 'Idiomas')
        c.setFont('Helvetica', 9)
        for lang in self.lang.items:
            lang_line = f'• {lang["Idioma"]} ({lang["Nível"]})'
            wrapped_lang = wrap_text(lang_line, max_label_width, 'Helvetica', 9)
            for part in wrapped_lang:
                y -= 12
                c.drawString(0.5*cm, y, part)

        # ====== ÁREA PRINCIPAL ======
        x0 = sidebar_w + margin
        y0 = h - margin

        c.setFont('Helvetica-Bold', 16)
        c.setFillColor(colors.HexColor("#1F4E79"))
        title_lines = wrap_text(data['Cargo / Título'], available_width, 'Helvetica-Bold', 16)
        for line in title_lines:
            c.drawString(x0, y0, line)
            y0 -= 18

        # Resumo Profissional (com ícone e divisor pontilhado)
        y0 -= 30
        # Desenha ícone de Resumo
        c.drawImage(ICONS['Resumo'], x0, y0-ICON_SIZE+4, width=ICON_SIZE, height=ICON_SIZE, preserveAspectRatio=True, mask='auto')
        # Título em negrito
        c.setFont('Helvetica-Bold', 12)
        c.drawString(x0 + ICON_SIZE + 6, y0, 'Resumo Profissional')
        # Divisor pontilhado abaixo
        y0 -= 6
        c.setLineWidth(0.4)
        c.setDash(2, 3)  
        c.line(x0, y0, x0 + available_width, y0)
        c.setDash()  # volta ao sólido
        c.setFillColor(colors.black)
        y0 -= 14
        text_obj = c.beginText(x0+10, y0)
        text_obj.setFont('Helvetica', 10)
        for paragraph in data['Resumo Profissional'].split('\n'):
            lines = wrap_text(paragraph, available_width, 'Helvetica', 10)
            for ln in lines:
                text_obj.textLine(ln)
            text_obj.textLine("")
        y0 = text_obj.getY()
        c.drawText(text_obj)

        # Experiência Profissional (com ícone e divisor pontilhado)
        y0 = text_obj.getY() - 20
        # Desenha ícone de Experiência
        c.drawImage(ICONS['Experiencia'], x0, y0-ICON_SIZE+4, width=ICON_SIZE, height=ICON_SIZE, preserveAspectRatio=True, mask='auto')
        c.setFillColor(colors.HexColor("#1F4E79"))
        c.setFont('Helvetica-Bold', 12)
        c.drawString(x0 + ICON_SIZE + 6, y0, 'Experiência Profissional')
        # Divisor pontilhado
        y0 -= 6
        c.setLineWidth(0.4)
        c.setDash(2, 3)
        c.line(x0, y0, x0 + available_width, y0)
        c.setDash()
        c.setFillColor(colors.black)
        y0 -= 16
        text_obj = c.beginText(x0+10, y0)
        text_obj.setFont('Helvetica', 10)
        for item in reversed(self.exp.items):
            # Primeira linha: Cargo – Empresa
            title_line = f"{item['Cargo']} – {item['Empresa']}"
            title_lines = wrap_text(title_line, available_width-10, 'Helvetica-Bold', 10)
            text_obj.setFont('Helvetica-Bold', 10)
            for tl in title_lines:
                text_obj.textLine(tl)

            # Segunda linha: Período
            text_obj.setFont('Helvetica-Oblique', 9)
            text_obj.textLine(f"    {item['Período']}")
            text_obj.setFont('Helvetica', 10)
            # descrição
            text_obj.setFont('Helvetica', 10)
            for paragraph in item['Descrição'].split('\n'):
                desc_lines = wrap_text(paragraph, available_width-20, 'Helvetica', 10)
                for i, dl in enumerate(desc_lines):
                    bullet = "• " if i == 0 else "  "
                    text_obj.textLine("    " + bullet + dl)
            text_obj.textLine("")
        y0 = text_obj.getY()
        c.drawText(text_obj)

        # Formação Acadêmica (com ícone e divisor pontilhado)
        y0 = text_obj.getY() - 20
        # Desenha ícone de Formação
        c.drawImage(ICONS['Formacao'], x0, y0-ICON_SIZE+4, width=ICON_SIZE, height=ICON_SIZE, preserveAspectRatio=True, mask='auto')
        c.setFillColor(colors.HexColor("#1F4E79"))
        c.setFont('Helvetica-Bold', 12)
        c.drawString(x0 + ICON_SIZE + 6, y0, 'Formação Acadêmica')
        # Divisor pontilhado
        y0 -= 6
        c.setLineWidth(0.4)
        c.setDash(2, 3)
        c.line(x0, y0, x0 + available_width, y0)
        c.setDash()
        c.setFillColor(colors.black)
        y0 -= 16
        text_obj = c.beginText(x0+10, y0)
        for item in reversed (self.edu.items):
            # Primeira linha: Curso – Escola
            title_line = f"{item['Curso']} – {item['Escola']}"
            title_lines = wrap_text(title_line, available_width-10, 'Helvetica-Bold', 10)
            text_obj.setFont('Helvetica-Bold', 10)
            for tl in title_lines:
                text_obj.textLine(tl)

            # Segunda linha: Período
            text_obj.setFont('Helvetica-Oblique', 9)
            text_obj.textLine(f"    {item['Período']}")
            text_obj.setFont('Helvetica', 10)
            text_obj.setFont('Helvetica', 10)
            for paragraph in item['Descrição'].split('\n'):
                desc_lines = wrap_text(paragraph, available_width-20, 'Helvetica', 10)
                for i, dl in enumerate(desc_lines):
                    bullet = "• " if i == 0 else "  "
                    text_obj.textLine("    " + bullet + dl)
            text_obj.textLine("")
        c.drawText(text_obj)

        # Salva PDF
        c.save()
        messagebox.showinfo('Sucesso', f'Currículo salvo em:\n{path}')

    def _generate_cover_letter(self):
        # coleta dados pessoais para cabeçalho
        nome     = self.entries['Nome Completo'].get().strip()
        cargo    = self.entries['Cargo / Título'].get().strip()
        email    = self.entries['E-mail'].get().strip()
        telefone = self.entries['Telefone'].get().strip()

        # coleta dados da carta
        data_entries = {k: e.get().strip() for k, e in self.cover_entries.items()}
        corpo     = self.cover_body.get('1.0', 'end').strip()
        despedida = self.cover_closing.get('1.0', 'end').strip()

        path = filedialog.asksaveasfilename(defaultextension='.pdf',
                                            filetypes=[('PDF files','*.pdf')],
                                            title='Salvar Carta de Apresentação')
        if not path:
            return

        doc = SimpleDocTemplate(path, pagesize=A4,
                                leftMargin=2*cm, rightMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY, leading=16))
        styles.add(ParagraphStyle(name='Right', alignment=TA_RIGHT))
        styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER, fontSize=14, spaceAfter=12))

        story = []
        # Cabeçalho centralizado com dados pessoais
        story.append(Paragraph(f"<b>{nome}</b> – {cargo}", styles['Center']))
        story.append(Paragraph(f"{email} | {telefone}", styles['Center']))
        story.append(Spacer(1, 12))

        # Data alinhada à direita
        story.append(Paragraph(data_entries['Data'], styles['Right']))
        story.append(Spacer(1, 12))

        # Endereço do destinatário
        addr = f"{data_entries['Destinatário']}<br/>{data_entries['Empresa']}"
        story.append(Paragraph(addr, styles['Normal']))
        story.append(Spacer(1, 12))

        # Saudação
        story.append(Paragraph(f"<b>{data_entries['Saudação']}</b>", styles['Normal']))
        story.append(Spacer(1, 12))

        # Corpo do texto
        for paragraph in corpo.split('\n\n'):
            story.append(Paragraph(paragraph.strip(), styles['Justify']))
            story.append(Spacer(1, 12))

        # Despedida e assinatura
        story.append(Paragraph(despedida.replace('\n', '<br/>'), styles['Normal']))
        story.append(Spacer(1, 24))
        story.append(Paragraph('__________________________', styles['Center']))
        story.append(Paragraph(nome, styles['Center']))

        doc.build(story)
        messagebox.showinfo('Sucesso', f'Carta salva em:\n{path}')

if __name__ == '__main__':
    CVGeneratorApp().mainloop()
