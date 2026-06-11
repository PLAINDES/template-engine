# app/services/html_converter.py
import mammoth
import re
from io import BytesIO

# Regex para variables normales: [VARIABLE]
VAR_PATTERN = re.compile(
    r'\[([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)\]',
    re.UNICODE,
)

# Regex para variables con formato HTML interno:
# [<strong>VAR</strong>], [<em>VAR</em>], [<strong><em>VAR</em></strong>], etc.
VAR_FORMATTED_PATTERN = re.compile(
    r'\[(?:<[^>]+>)*'                          # corchete + tags de apertura opcionales
    r'([A-ZÁÉÍÓÚÑÜa-záéíóúñü]'               # primer char de la variable
    r'[A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)'          # resto de la variable
    r'(?:</[^>]+>)*\]',                        # tags de cierre + corchete
    re.UNICODE,
)

# Regex para variables con nombres que contienen espacios o puntos:
# [NUMERO DE ALTERNATIVAS_TECNICAS], [TIPO_DE MONEDA], [FECHA_R.M], [CRITERIO_3.1]
VAR_SPECIAL_CHARS = re.compile(
    r'\[([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_\s\.]*[A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_])\]',
    re.UNICODE,
)

# Regex para variables donde el corchete de cierre está fuera del tag:
# [DESCRIPCION_SOSTENIBILIDAD_AMBIENTAL<strong>].</strong>
VAR_CLOSE_OUTSIDE = re.compile(
    r'\[([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)(?:<[^>]+>)+\]',
    re.UNICODE,
)

# Regex para variables donde los corchetes están en su propio <strong> separado:
# <strong>[</strong>VARIABLE<strong>]</strong>
# Cubre: [ORIGEN_INGRESOS], [DESCRIPCION_SPE], [DESCRIPCION_DURACION_PROYECTO],
#        [DETALLES_ADICIONALES_...], [BRECHA_CUANTITATIVA_CUALITATIVA]
VAR_BRACKET_STRONG = re.compile(
    r'<strong>\[</strong>'
    r'([A-ZÁÉÍÓÚÑÜa-záéíóúñü][A-ZÁÉÍÓÚÑÜa-záéíóúñü0-9_]*)'
    #r'<strong>\]</strong>',
    r'<strong>\][^<]*</strong>',
    re.UNICODE,
)

# Eliminar TOC
TOC_PARA_RE   = re.compile(
    r'<p[^>]*>\s*(?:<a\s+href="#_Toc[^"]*"[^>]*>.*?</a>\s*)+</p>',
    re.DOTALL,
)
TOC_LINK_RE   = re.compile(r'<a\s+href="#_Toc[^"]*"[^>]*>(.*?)</a>', re.DOTALL)
TOC_ANCHOR_RE = re.compile(r'<a\s+id="_Toc[^"]*"[^>]*>\s*</a>', re.DOTALL)


def _mark_variable(key: str) -> str:
    return f'<span data-var="{key}" class="doc-var">[{key}]</span>'


def _fix_formatted_vars(html: str) -> str:
    """
    Convierte variables al span editable, en este orden de especificidad:

    1. VAR_BRACKET_STRONG    -> <strong>[</strong>VAR<strong>]</strong>
    2. VAR_CLOSE_OUTSIDE     -> [VAR<strong>]</strong>
    3. VAR_FORMATTED_PATTERN -> [<strong>VAR</strong>]
    4. VAR_SPECIAL_CHARS     -> [NOMBRE CON ESPACIOS] / [FECHA_R.M]
    5. VAR_PATTERN           -> [VARIABLE] normales
    """
    # 1. Corchetes en su propio <strong> separado del nombre
    html = VAR_BRACKET_STRONG.sub(
        lambda m: _mark_variable(m.group(1)),
        html,
    )
    # 2. Corchete de cierre fuera del tag
    html = VAR_CLOSE_OUTSIDE.sub(
        lambda m: _mark_variable(m.group(1)),
        html,
    )
    # 3. Variable con tags internos de formato
    html = VAR_FORMATTED_PATTERN.sub(
        lambda m: _mark_variable(m.group(1)),
        html,
    )
    # 4. Variables con espacios o puntos en el nombre
    html = VAR_SPECIAL_CHARS.sub(
        lambda m: _mark_variable(m.group(1)),
        html,
    )
    # 5. Variables normales que quedaron sin marcar
    html = VAR_PATTERN.sub(
        lambda m: _mark_variable(m.group(1)),
        html,
    )
    return html


def _remove_toc_from_html(html: str) -> str:
    html = TOC_PARA_RE.sub("", html)
    html = TOC_LINK_RE.sub(r"\1", html)
    html = TOC_ANCHOR_RE.sub("", html)
    return html


def docx_to_html(docx_buffer: bytes) -> dict:
    style_map = """
        p[style-name='Heading 1'] => h1.doc-h1:fresh
        p[style-name='Heading 2'] => h2.doc-h2:fresh
        p[style-name='Heading 3'] => h3.doc-h3:fresh
        p[style-name='Heading 4'] => h4.doc-h4:fresh
        p[style-name='Titulo 1']  => h1.doc-h1:fresh
        p[style-name='Titulo 2']  => h2.doc-h2:fresh
        p[style-name='Titulo 3']  => h3.doc-h3:fresh
        p[style-name='heading 1'] => h1.doc-h1:fresh
        p[style-name='heading 2'] => h2.doc-h2:fresh
        p[style-name='heading 3'] => h3.doc-h3:fresh
        p[style-name='Normal']    => p.doc-p:fresh
        p[style-name='Body Text'] => p.doc-p:fresh
        p[style-name='toc 1']     => p.toc-entry:fresh
        p[style-name='toc 2']     => p.toc-entry:fresh
        p[style-name='toc 3']     => p.toc-entry:fresh
        p[style-name='toc 4']     => p.toc-entry:fresh
        p[style-name='toc 5']     => p.toc-entry:fresh
        r[style-name='Strong']    => strong
        r[style-name='Emphasis']  => em
    """

    result = mammoth.convert_to_html(
        BytesIO(docx_buffer),
        style_map=style_map,
        convert_image=mammoth.images.img_element(
            lambda image: {
                "src": "data:{};base64,{}".format(
                    image.content_type,
                    __import__("base64").b64encode(image.read()).decode("utf-8"),
                )
            }
        ),
    )

    html = result.value

    # 1. Eliminar TOC
    html = _remove_toc_from_html(html)

    # 2. Eliminar párrafos toc-entry
    html = re.sub(
        r'<p[^>]*class="toc-entry"[^>]*>.*?</p>',
        "",
        html,
        flags=re.DOTALL,
    )

    # 3. Marcar variables — todos los patrones en orden de especificidad
    html = _fix_formatted_vars(html)

    return {
        "html":     html,
        "messages": [str(m) for m in result.messages],
    }