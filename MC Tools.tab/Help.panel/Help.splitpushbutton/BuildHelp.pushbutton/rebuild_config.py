#! python3
r"""
Build Help (script.py)  -  Murray offline help pages for pyRevit
Python 3, standard library only.

FOLDER LAYOUT (everything lives in one pyRevit button):
    MurrayHelpDocuments.pushbutton\
        help.md            home page (welcome text, What's New)
        help.html          generated
        script.py, bundle.yaml, icon.png
        topics\
            PipeLabels\
                help.md    one folder per topic
                images\    screenshots for that topic
                files\     attachments (PDFs etc.)
                help.html  generated

ADD A TOPIC:
    Create topics\NewTopicName\help.md (first line "# Title"), add images\ if needed,
    then run the build. The sidebar picks it up automatically.

BUILD (every time you edit a page):
    Click the Build Help button in Revit.
    Outside Revit: python script.py build "path\to\Help.panel"

IMPORT (one time, to bring Notion pages over):
    python script.py import "ExportBlock-xxxx.zip" "path\to\MurrayHelpDocuments.pushbutton"
    Creates topics\<Name>\ for each Notion page. Existing help.md files are never overwritten.

LOGO (optional):
    Put murray_logo.png in the same folder as this script. The build copies it into
    MurrayHelpDocuments.pushbutton and shows it in the yellow header. No file = no logo.

F1 HELP ON A TOOL BUTTON:
    Put a small text file named topic.txt inside the button folder (next to its script).
    (Do not name it help.anything: pyRevit opens help.* files when F1 is pressed.)
    It holds the topic folder name, one per line, for example:
        MCPointloadPipe
    The build writes a help.html into that button folder that opens the topic, so F1 on the
    button shows it. The link belongs to the button: delete, rename or move the button folder
    and its help link goes with it. Remove topic.txt and the next build removes the link.
    A button covered by two topics lists both names and F1 shows a short chooser page.
    A button folder that already has its own hand-made help.html is never overwritten.

OPTIONAL HEADER LINES (at the very top of a topic help.md, above the "# Title" line):
    summary: One sentence shown on the What's New card.
    date: 2026-10-16
        Only needed to override the automatic date. By default the build records the date
        a topic was first built or last changed in topics.manifest.json (keep that file).

WHAT'S NEW ON THE HOME PAGE:
    Put the single line {{whats-new}} under the "## What's New" heading in the home help.md.
    The list is generated from the topics. Each reader's browser remembers which topics
    they have seen, so NEW and IMPROVED badges are per person. Editing a topic's text or
    images marks it IMPROVED; changing only the header lines above does not.

MARKDOWN SUPPORTED IN help.md
    # Page title                (first line, used as the page and button title)
    ## Section heading
    ### Sub heading
    - bullet item               (indent 4 spaces per level to nest)
    1. numbered item
    > Note text                 (yellow highlighted note; can be indented under a list item)
    ![description](images/file.png)   (on its own line; indent to keep it under a list item)
    **bold**   *italic*   ~~strikethrough~~   `code`
    [link text](https://...)    or just paste a bare https:// address
    [Other page](../OtherTopic/help.html)             (link to another topic, from a topic)
    [Catalog PDF](files/catalog.pdf)                   (link to a file in this topic)
    End a line with a single backslash character to force a line break.
    ---                         (horizontal line)
    Plain lines become paragraphs. Blank lines separate paragraphs.
"""
import base64
import datetime
import hashlib
import html
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
from html.parser import HTMLParser
from urllib.parse import quote, unquote

# --------------------------------------------------------------------------
# Page template
# --------------------------------------------------------------------------
CSS = """
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;background:#fff;color:#000;font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.55}
.band{background:#FFE716;border-bottom:2px solid #000}
.band-inner{display:flex;align-items:center;gap:20px;padding:16px 28px}
.band-logo{height:56px;width:auto;flex:none}
.band small{display:block;font-size:14px;font-weight:bold;margin-bottom:4px}
.band small a{text-decoration:none}
.band small a:hover{text-decoration:underline}
.band h1{margin:0;font-size:30px;line-height:1.2}
.layout{display:flex;align-items:flex-start}
.side{flex:none;width:300px;position:sticky;top:0;max-height:100vh;overflow:auto;padding:18px 0 24px}
.side input{display:block;width:calc(100% - 40px);margin:0 20px 12px;padding:7px 10px;font:inherit;font-size:14px;border:1px solid #000;border-radius:0}
.side ul{list-style:none;margin:0;padding:0}
.side li{margin:0}
.side a{display:block;padding:7px 20px 7px 16px;border-left:4px solid transparent;text-decoration:none;font-size:15px;line-height:1.3}
.side a:hover{border-left-color:#000;text-decoration:underline}
.side a[aria-current="page"]{background:#FFE716;border-left-color:#000;font-weight:bold}
.side .home-link{font-weight:bold;border-bottom:1px solid #000;margin-bottom:6px;padding-bottom:10px}
.side .none{display:none;padding:7px 20px;font-size:14px}
.content{flex:1;min-width:0;min-height:100vh;padding:24px 36px 0;border-left:1px solid #000;margin-left:-1px}
.content>.page{max-width:900px}
h2{font-size:22px;margin:32px 0 10px;padding-bottom:4px;border-bottom:2px solid #FFE716;line-height:1.25}
h3{font-size:18px;margin:24px 0 8px;line-height:1.25}
p{margin:8px 0}
ul,ol{margin:4px 0;padding-left:26px}
li{margin:3px 0}
.page>ul>li,.page>ol>li{margin-top:12px}
a{color:#000;text-decoration:underline}
a:focus-visible,img:focus-visible,button:focus-visible,textarea:focus-visible,input:focus-visible{outline:3px solid #000;outline-offset:2px}
figure{margin:12px 0}
figure img{display:block;max-width:100%;height:auto;border:1px solid #000;cursor:zoom-in}
.note{background:#FFE716;padding:6px 12px;margin:6px 0;border-left:4px solid #000}
.missing{border:2px dashed #000;padding:8px 12px;font-weight:bold}
code{font-family:Consolas,monospace;font-size:14px;border:1px solid #000;padding:0 4px}
hr{border:0;border-top:1px solid #000;margin:20px 0}
.home{display:grid;grid-template-columns:minmax(0,1fr) 320px;gap:40px;max-width:1200px}
.home>.page>h2:first-child{margin-top:8px}
.feedback{border:1px solid #000;padding:4px 20px 20px;align-self:start;position:sticky;top:24px}
.feedback h2{margin-top:16px}
.feedback textarea{display:block;width:100%;min-height:140px;margin:10px 0;padding:8px;font:inherit;font-size:15px;border:1px solid #000;border-radius:0;resize:vertical}
.feedback button{font:inherit;font-weight:bold;background:#FFE716;color:#000;border:2px solid #000;padding:8px 22px;cursor:pointer}
.feedback button:hover{background:#000;color:#fff}
.feedback .small{font-size:13px}
.lightbox{position:fixed;inset:0;background:rgba(0,0,0,.88);display:none;align-items:center;justify-content:center;padding:24px;cursor:zoom-out;z-index:10}
.lightbox.open{display:flex}
.lightbox img{max-width:100%;max-height:100%;background:#fff}
.wn{margin:12px 0 8px}
.wn-head{font-weight:bold;margin:16px 0 4px}
.wn-card{border:1px solid #000;border-top:4px solid #FFE716;padding:4px 16px 10px;margin:10px 0}
.wn-card h3{margin:8px 0 2px;font-size:17px}
.wn-card .meta{font-size:13px;margin:0 0 4px}
.wn-card p{margin:4px 0}
.wn-note{font-size:13px}
.badge{display:inline-block;font-size:12px;font-weight:bold;padding:1px 8px;margin-left:8px;border:1px solid #000;background:#FFE716;color:#000;vertical-align:middle}
.badge.upd{background:#fff}
.btn{font:inherit;font-weight:bold;background:#FFE716;color:#000;border:2px solid #000;padding:6px 16px;margin:8px 10px 6px 0;cursor:pointer}
.btn:hover{background:#000;color:#fff}
footer{max-width:900px;margin:48px 0 0;padding:12px 0 28px;font-size:12px;border-top:1px solid #000}
@media (max-width:1100px){.home{grid-template-columns:1fr}.feedback{position:static}}
@media (max-width:760px){.layout{display:block}.side{position:static;width:auto;max-height:none;border-bottom:1px solid #000}.content{padding:20px;border-left:0;margin-left:0;min-height:0}}
@media print{.side,.feedback,.lightbox{display:none!important}figure img{cursor:auto}.content{padding:0}}
"""

JS = """
(function(){
  var box=document.getElementById('lb'), big=box.querySelector('img');
  document.querySelectorAll('figure img').forEach(function(im){
    im.tabIndex=0;
    function open(){big.src=im.src;box.classList.add('open');}
    im.addEventListener('click',open);
    im.addEventListener('keydown',function(e){if(e.key==='Enter'){open();}});
  });
  function close(){box.classList.remove('open');big.removeAttribute('src');}
  box.addEventListener('click',close);
  document.addEventListener('keydown',function(e){if(e.key==='Escape'){close();}});

  var f=document.getElementById('topic-filter');
  if(f){
    var items=document.querySelectorAll('#topics li'), none=document.getElementById('no-topics');
    f.addEventListener('input',function(){
      var q=f.value.toLowerCase().trim(), shown=0;
      items.forEach(function(li){var hit=li.textContent.toLowerCase().indexOf(q)>-1;li.style.display=hit?'':'none';if(hit)shown++;});
      none.style.display=shown?'none':'block';
    });
  }
  var cur=document.querySelector('.side a[aria-current="page"]');
  if(cur&&cur.scrollIntoView){cur.scrollIntoView({block:'nearest'});}

  var send=document.getElementById('fb-send');
  if(send){
    send.addEventListener('click',function(){
      var t=document.getElementById('fb-text'), msg=t.value.trim();
      if(!msg){t.focus();return;}
      window.location.href='mailto:'+send.getAttribute('data-to')
        +'?cc='+encodeURIComponent(send.getAttribute('data-cc'))
        +'&subject='+encodeURIComponent('Murray Help feedback')
        +'&body='+encodeURIComponent(msg);
    });
  }
})();
"""

JS_SEEN = r"""
(function(){
  var KEY='murrayHelpSeen_v1';
  function load(){
    var r;
    try{r=window.localStorage.getItem(KEY);}catch(e){return null;}
    if(!r){return {};}
    try{var o=JSON.parse(r);return (o&&typeof o==='object')?o:{};}catch(e2){return {};}
  }
  function save(s){try{window.localStorage.setItem(KEY,JSON.stringify(s));return true;}catch(e){return false;}}

  var idm=document.querySelector('meta[name="murray-topic-id"]');
  var hm=document.querySelector('meta[name="murray-content-hash"]');
  if(idm&&hm){var s0=load();if(s0){s0[idm.content]=hm.content;save(s0);}}

  var box=document.getElementById('whats-new'), dm=document.getElementById('topics-data');
  if(!box||!dm){return;}
  var data=JSON.parse(dm.textContent);
  var seen=load(), canSave=(seen!==null);
  if(seen===null){seen={};}
  var showAll=!canSave;
  var MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  function el(tag,cls,text){var e=document.createElement(tag);if(cls){e.className=cls;}if(text!==undefined){e.textContent=text;}return e;}
  function fmt(d){var p=(d||'').split('-');if(p.length!==3){return d||'';}return MON[parseInt(p[1],10)-1]+' '+parseInt(p[2],10)+', '+p[0];}
  function status(t){if(!(t.id in seen)){return 'new';}if(seen[t.id]!==t.hash){return 'upd';}return 'seen';}
  function card(t,st){
    var c=el('div','wn-card'), h=el('h3'), a=el('a',null,t.title);
    a.href=t.href;h.appendChild(a);
    if(st==='new'){h.appendChild(el('span','badge','NEW'));}
    else if(st==='upd'){h.appendChild(el('span','badge upd','IMPROVED'));}
    c.appendChild(h);
    c.appendChild(el('div','meta',[fmt(t.date),t.type].filter(Boolean).join(' \u00b7 ')));
    if(t.summary){c.appendChild(el('p',null,t.summary));}
    return c;
  }
  function render(){
    box.textContent='';
    var items=data.slice().sort(function(a,b){return a.date<b.date?1:(a.date>b.date?-1:0);});
    var fresh=[], upd=[];
    items.forEach(function(t){var s=status(t);if(s==='new'){fresh.push(t);}else if(s==='upd'){upd.push(t);}});
    var anyNew=(fresh.length+upd.length)>0;
    if(!canSave){box.appendChild(el('p','wn-note','This browser is not saving which topics you have read, so every topic is listed.'));}
    if(!showAll){
      if(!anyNew){
        box.appendChild(el('p','wn-head','You are all caught up.'));
        box.appendChild(el('p',null,'No new or improved topics since you last checked.'));
      }else{
        if(fresh.length){box.appendChild(el('p','wn-head','New ('+fresh.length+')'));fresh.forEach(function(t){box.appendChild(card(t,'new'));});}
        if(upd.length){box.appendChild(el('p','wn-head','Improved ('+upd.length+')'));upd.forEach(function(t){box.appendChild(card(t,'upd'));});}
      }
    }else{
      box.appendChild(el('p','wn-head','All topics ('+items.length+')'));
      items.forEach(function(t){box.appendChild(card(t,status(t)));});
    }
    var bar=el('div');
    if(canSave&&anyNew){
      var b1=el('button','btn','Mark all as read');b1.type='button';
      b1.onclick=function(){data.forEach(function(t){seen[t.id]=t.hash;});save(seen);showAll=false;render();};
      bar.appendChild(b1);
    }
    if(canSave){
      var b2=el('button','btn',showAll?'Show only new':'Show all topics');b2.type='button';
      b2.onclick=function(){showAll=!showAll;render();};
      bar.appendChild(b2);
    }
    box.appendChild(bar);
  }
  render();
  window.addEventListener('pageshow',function(e){if(e.persisted){var s=load();if(s){seen=s;}render();}});
})();
"""

TEMPLATE = """<!DOCTYPE html>
<!-- GENERATED FILE. Do not edit. Edit help.md and click the Build Help button. -->
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
{meta}
<style>{css}</style>
</head>
<body>
<header class="band"><div class="band-inner">{logo}<div><small>{home}</small><h1>{title}</h1></div></div></header>
<div class="layout">
{sidebar}
<main class="content">
{body}
<footer>Last built {date}. </footer>
</main>
</div>
<div class="lightbox" id="lb" role="dialog" aria-label="Enlarged image"><img alt=""></div>
<script>{js}</script>
</body>
</html>
"""

FEEDBACK = """<aside class="feedback" aria-labelledby="fb-h">
<h2 id="fb-h">Feedback and Suggestions</h2>
<p>Found a mistake, or have an idea for a new help topic? Let us know.</p>
<label for="fb-text" class="small">Your message</label>
<textarea id="fb-text"></textarea>
<button type="button" id="fb-send" data-to="{email}" data-cc="{cc}">Send</button>
<p class="small">Send opens a new email in Outlook with your message filled in. Review it and click Send there.
If nothing opens, email <a href="mailto:{email}">{email}</a> directly.</p>
</aside>"""

# Where the feedback box on the home page sends to.
FEEDBACK_EMAIL = "moatman@murraycompany.com"
FEEDBACK_CC = "fcruz@murraycompany.com"

# Bundle folder that holds the whole help site, and its topics subfolder.
INDEX_FOLDER = "MurrayHelpDocuments.pushbutton"
TOPICS_DIR = "topics"

# --------------------------------------------------------------------------
# Markdown -> HTML
# --------------------------------------------------------------------------
RE_ITEM = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
RE_IMG = re.compile(r"^(\s*)!\[([^\]]*)\]\(([^)]+)\)\s*$")
RE_NOTE = re.compile(r"^(\s*)>\s?(.*)$")
RE_HEAD = re.compile(r"^(#{1,3})\s+(.*)$")


def inline(text, page_dir=None, warnings=None):
    text = html.escape(text, quote=False)
    codes = []

    def keep(m):
        codes.append("<code>%s</code>" % m.group(1))
        return "\x00%d\x00" % (len(codes) - 1)
    text = re.sub(r"`([^`]+)`", keep, text)

    def link(m):
        label, href = m.group(1), m.group(2)
        if href.startswith(("http://", "https://", "mailto:")):
            return '<a href="%s" target="_blank" rel="noopener">%s</a>' % (href.replace('"', "%22"), label)
        if page_dir is not None and warnings is not None:
            target = os.path.join(page_dir, unquote(html.unescape(href).split("#")[0]))
            sibling_md = target.endswith("help.html") and os.path.exists(target[:-4] + "md")
            if not os.path.exists(target) and not sibling_md:
                warnings.append("Link target not found: " + html.unescape(href))
        return '<a href="%s">%s</a>' % (href.replace('"', "%22"), label)
    def link_keep(m):
        codes.append(link(m))
        return "\x00%d\x00" % (len(codes) - 1)
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", link_keep, text)
    text = re.sub(r"(?<![\w\x00])(https?://[^\s<]+[^\s<.,;:!?)])",
                  lambda m: '<a href="%s" target="_blank" rel="noopener">%s</a>' % (m.group(1), m.group(1)), text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<em>\1</em>", text)
    text = re.sub(r"~~(.+?)~~", r"<del>\1</del>", text)
    text = re.sub(r"\x00(\d+)\x00", lambda m: codes[int(m.group(1))], text)
    return text


def indent_of(s):
    return len(s.expandtabs(4))


def md_to_html(md, page_dir, warnings):
    lines = md.replace("\r\n", "\n").split("\n")
    title = "Help"
    out = []
    stack = []  # open lists: [indent, tag]
    last_para = None  # (indent, index in out) for joining wrapped lines

    def close_until(pred):
        while stack and pred(stack[-1]):
            out.append("</li></%s>" % stack.pop()[1])

    first_heading_done = False
    for raw in lines:
        if not raw.strip():
            last_para = None
            continue
        m = RE_HEAD.match(raw)
        if m and not raw.startswith(" "):
            close_until(lambda s: True)
            level, text = len(m.group(1)), m.group(2).strip()
            if level == 1 and not first_heading_done:
                title = text
                first_heading_done = True
            else:
                tag = "h%d" % max(level, 2)
                out.append("<%s>%s</%s>" % (tag, inline(text, page_dir, warnings), tag))
            last_para = None
            continue
        if raw.strip() == "---":
            close_until(lambda s: True)
            out.append("<hr>")
            last_para = None
            continue

        if raw.strip() == "{{whats-new}}":
            close_until(lambda s: True)
            out.append('<div id="whats-new" class="wn"></div>')
            last_para = None
            continue

        m = RE_ITEM.match(raw)
        if m:
            ind = indent_of(m.group(1))
            tag = "ol" if m.group(2)[0].isdigit() else "ul"
            close_until(lambda s: s[0] > ind)
            if stack and stack[-1][0] == ind and stack[-1][1] != tag:
                close_until(lambda s: s[0] >= ind)
            if stack and stack[-1][0] == ind:
                out.append("</li><li>")
            else:
                start = ""
                if tag == "ol":
                    n = int(m.group(2)[:-1])
                    if n != 1:
                        start = ' start="%d"' % n
                out.append("<%s%s><li>" % (tag, start))
                stack.append([ind, tag])
            out.append(inline(m.group(3).strip(), page_dir, warnings))
            last_para = None
            continue

        m = RE_IMG.match(raw)
        if m:
            ind = indent_of(m.group(1))
            close_until(lambda s: s[0] >= ind)
            alt, src = m.group(2), m.group(3).strip()
            if not src.startswith(("http://", "https://")) and not os.path.isfile(
                    os.path.join(page_dir, unquote(src))):
                warnings.append("Image not found: " + src)
                out.append('<p class="missing">Missing image: %s</p>' % html.escape(src))
            else:
                out.append('<figure><img src="%s" alt="%s"></figure>'
                           % (html.escape(src), html.escape(alt)))
            last_para = None
            continue

        m = RE_NOTE.match(raw)
        if m:
            ind = indent_of(m.group(1))
            close_until(lambda s: s[0] >= ind)
            out.append('<div class="note">%s</div>' % inline(m.group(2).strip(), page_dir, warnings))
            last_para = None
            continue

        # Paragraph (joins with the previous line if no blank line between)
        ind = indent_of(raw[:len(raw) - len(raw.lstrip())])
        txt = raw.strip()
        brk = txt.endswith("\\")
        if brk:
            txt = txt[:-1].rstrip()
        piece = inline(txt, page_dir, warnings) + ("<br>" if brk else "")
        if last_para and last_para[0] == ind:
            i = last_para[1]
            prev = out[i][:-4]
            out[i] = prev + ("" if prev.endswith("<br>") else " ") + piece + "</p>"
            continue
        close_until(lambda s: s[0] >= ind)
        out.append("<p>%s</p>" % piece)
        last_para = (ind, len(out) - 1)

    close_until(lambda s: True)
    return title, "\n".join(out)


def md_title(md_path):
    with open(md_path, encoding="utf-8") as fh:
        for line in fh:
            m = re.match(r"#\s+(.*)", line)
            if m:
                return m.group(1).strip()
    return os.path.basename(os.path.dirname(md_path))


def find_help_root(folder):
    """Accepts the help bundle itself, its parent panel, or anything above it."""
    folder = os.path.abspath(folder)
    if os.path.basename(folder) == INDEX_FOLDER:
        return folder
    for root, dirs, _ in os.walk(folder):
        if INDEX_FOLDER in dirs:
            return os.path.join(root, INDEX_FOLDER)
    return None


def collect_topics(help_root):
    tdir = os.path.join(help_root, TOPICS_DIR)
    topics = []
    if os.path.isdir(tdir):
        for name in os.listdir(tdir):
            md = os.path.join(tdir, name, "help.md")
            if os.path.isfile(md):
                topics.append((name, md_title(md)))
    return sorted(topics, key=lambda t: t[1].lower())


# --------------------------------------------------------------------------
# Topic header lines, content hash, manifest
# --------------------------------------------------------------------------
RE_META = re.compile(r"^(date|buttons|summary)\s*:\s*(.*)$", re.I)
TOPIC_FILE = "topic.txt"
LEGACY_TOPIC_FILE = "help.topic"   # renamed automatically: pyRevit treats help.* files as the F1 file
MANIFEST_FILE = "topics.manifest.json"
STUB_MARK = "murray-help-stub"
SKIP_DIRS = ("__pycache__", "lib", "hooks", "images", "files", "topics")


def split_header(raw):
    """Returns (meta dict, Markdown body). Header lines are key: value lines at the top."""
    lines = raw.replace("\r\n", "\n").split("\n")
    meta, i = {}, 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        m = RE_META.match(s)
        if not m:
            break
        meta[m.group(1).lower()] = m.group(2).strip()
        i += 1
    return meta, "\n".join(lines[i:])


def norm_name(name):
    return re.sub(r"[\s_]+", " ", name.strip().lower()).strip()


def content_hash(page_dir, body):
    h = hashlib.sha1()
    h.update(body.replace("\r\n", "\n").strip().encode("utf-8"))
    for sub in ("images", "files"):
        d = os.path.join(page_dir, sub)
        if not os.path.isdir(d):
            continue
        for name in sorted(os.listdir(d)):
            p = os.path.join(d, name)
            if os.path.isfile(p):
                h.update(("\0%s/%s\0" % (sub, name)).encode("utf-8"))
                with open(p, "rb") as fh:
                    for chunk in iter(lambda: fh.read(1048576), b""):
                        h.update(chunk)
    return h.hexdigest()[:16]


def load_manifest(root):
    path = os.path.join(root, MANIFEST_FILE)
    if not os.path.isfile(path):
        return {}, False
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        return data.get("topics", {}), True
    except (ValueError, OSError):
        print("   ! %s could not be read; dates restart from today" % MANIFEST_FILE)
        return {}, True


def save_manifest(root, topics):
    path = os.path.join(root, MANIFEST_FILE)
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        json.dump({"version": 1, "topics": topics}, fh, indent=2, sort_keys=True)
        fh.write("\n")


def gather_info(root, topics):
    """One pass over every topic: title, hash, date, buttons. Updates topics.manifest.json."""
    prev, had_manifest = load_manifest(root)
    today = datetime.date.today().isoformat()
    infos, manifest = {}, {}
    for folder, title in topics:
        md_path = os.path.join(root, TOPICS_DIR, folder, "help.md")
        with open(md_path, encoding="utf-8") as fh:
            meta, body = split_header(fh.read())
        digest = content_hash(os.path.dirname(md_path), body)
        old = prev.get(folder) or {}
        if old.get("hash") == digest and old.get("date"):
            date = old["date"]
        elif old or had_manifest:
            date = today
        else:
            date = datetime.date.fromtimestamp(os.path.getmtime(md_path)).isoformat()
        manifest[folder] = {"hash": digest, "date": date}
        shown = date
        if meta.get("date"):
            try:
                datetime.datetime.strptime(meta["date"], "%Y-%m-%d")
                shown = meta["date"]
            except ValueError:
                print("   ! %s: date line must look like 2026-10-16 (ignored)" % folder)
        if "buttons" in meta:
            print("   ! %s: the buttons: line is no longer used. Put a topic.txt file in each button folder." % folder)
        infos[folder] = {"id": folder, "title": title, "hash": digest, "date": shown,
                         "summary": meta.get("summary", ""), "type": "General"}
    save_manifest(root, manifest)
    return infos


def home_data_json(infos):
    items = [{"id": i["id"], "title": i["title"], "hash": i["hash"], "date": i["date"],
              "type": i["type"], "summary": i["summary"],
              "href": "%s/%s/help.html" % (TOPICS_DIR, quote(i["id"]))}
             for i in infos.values()]
    return json.dumps(items, ensure_ascii=True).replace("</", "<\\/")


# --------------------------------------------------------------------------
# F1 help links: small help.html files inside the tool button folders
# --------------------------------------------------------------------------
STUB_REDIRECT = """<!DOCTYPE html>
<!-- GENERATED STUB. Do not edit. Edit the buttons: line in the topic help.md, then click Build Help. -->
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="murray-help-stub" content="1">
<meta http-equiv="refresh" content="0; url={url}">
<title>{title}</title>
</head>
<body style="font-family:Arial,Helvetica,sans-serif">
<p>Opening <a href="{url}">{title}</a>.</p>
<script>location.replace({url_js});</script>
</body>
</html>
"""

STUB_CHOOSER = """<!DOCTYPE html>
<!-- GENERATED STUB. Do not edit. Edit the buttons: line in the topic help.md, then click Build Help. -->
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="murray-help-stub" content="1">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>{css}</style>
</head>
<body>
<header class="band"><div class="band-inner">{logo}<div><h1>{title}</h1></div></div></header>
<div class="layout"><main class="content" style="border-left:0;margin-left:0">
<div class="page"><p>More than one help topic covers this button. Pick one:</p>
<ul>
{items}
</ul></div>
</main></div>
</body>
</html>
"""


def find_extension_root(folder):
    p = os.path.abspath(folder)
    while True:
        if os.path.basename(p).lower().endswith(".extension"):
            return p
        parent = os.path.dirname(p)
        if parent == p:
            return None
        p = parent


def rel_url(from_dir, target):
    return quote(os.path.relpath(target, from_dir).replace(os.sep, "/"), safe="/")


def stub_html(bdir, folders, infos, root, logo):
    topic_files = [(infos[f]["title"], os.path.join(root, TOPICS_DIR, f, "help.html")) for f in folders]
    topic_files.sort(key=lambda t: t[0].lower())
    if len(topic_files) == 1:
        title, target = topic_files[0]
        url = rel_url(bdir, target)
        return STUB_REDIRECT.format(url=html.escape(url), url_js=json.dumps(url), title=html.escape(title))
    stem = os.path.splitext(os.path.basename(bdir))[0].replace("_", " ")
    logo_tag = ""
    if logo:
        logo_tag = '<img class="band-logo" src="%s" alt="Murray Company">' % html.escape(
            rel_url(bdir, os.path.join(root, "murray_logo.png")))
    items = "\n".join('<li><a href="%s">%s</a></li>' % (html.escape(rel_url(bdir, t)), html.escape(n))
                      for n, t in topic_files)
    return STUB_CHOOSER.format(title=html.escape("Help: " + stem), css=CSS, logo=logo_tag, items=items)


def scan_topic_files(root, topics):
    """Finds every topic.txt file in the extension.
    Returns (extension folder, {button folder: [topic folders]}, [folders that hold a help.html])."""
    ext = find_extension_root(root)
    if not ext:
        print("   ! Extension folder not found above %s; F1 help links skipped" % root)
        return None, {}, []
    lookup = {}
    for folder, title in topics:
        lookup[norm_name(folder)] = folder
        lookup[norm_name(title)] = folder
    claims, html_dirs = {}, []
    for dp, dns, fns in os.walk(ext):
        dns[:] = [d for d in dns if not d.startswith(".") and d.lower() not in SKIP_DIRS]
        lower = dict((f.lower(), f) for f in fns)
        if LEGACY_TOPIC_FILE in lower:
            old_path = os.path.join(dp, lower[LEGACY_TOPIC_FILE])
            if TOPIC_FILE in lower:
                print("   ! Both %s and %s exist in %s. Delete %s." % (
                    LEGACY_TOPIC_FILE, TOPIC_FILE, os.path.relpath(dp, ext), LEGACY_TOPIC_FILE))
            else:
                os.replace(old_path, os.path.join(dp, TOPIC_FILE))
                print("Renamed %s to %s in %s" % (LEGACY_TOPIC_FILE, TOPIC_FILE, os.path.relpath(dp, ext)))
                lower[TOPIC_FILE] = TOPIC_FILE
            lower.pop(LEGACY_TOPIC_FILE, None)
        if "help.html" in lower:
            html_dirs.append(dp)
        if TOPIC_FILE not in lower or os.path.abspath(dp) == os.path.abspath(root):
            continue
        with open(os.path.join(dp, lower[TOPIC_FILE]), encoding="utf-8-sig", errors="replace") as fh:
            lines = [ln.strip() for ln in fh.read().splitlines()]
        for ln in lines:
            if not ln or ln.startswith("#"):
                continue
            folder = lookup.get(norm_name(ln))
            if not folder:
                print("   ! Unknown topic '%s' in %s" % (ln, os.path.relpath(os.path.join(dp, lower[TOPIC_FILE]), ext)))
            elif folder not in claims.setdefault(dp, []):
                claims[dp].append(folder)
    return ext, claims, html_dirs


def write_stubs(root, infos, logo, ext, claims, html_dirs):
    if not ext:
        return
    written = same = removed = 0
    for bdir, folders in sorted(claims.items()):
        target = os.path.join(bdir, "help.html")
        content = stub_html(bdir, folders, infos, root, logo)
        if os.path.isfile(target):
            with open(target, encoding="utf-8", errors="replace") as fh:
                old = fh.read()
            if STUB_MARK not in old:
                print("   ! Not overwritten, has its own help.html: " + os.path.relpath(bdir, ext))
                continue
            if old == content:
                same += 1
                continue
        with open(target, "w", encoding="utf-8") as fh:
            fh.write(content)
        written += 1
    for d in html_dirs:
        target = os.path.join(d, "help.html")
        if d in claims or not os.path.isfile(target):
            continue
        with open(target, encoding="utf-8", errors="replace") as fh:
            is_stub = STUB_MARK in fh.read()
        if is_stub:
            os.remove(target)
            removed += 1
            print("Removed help link, no %s beside it: %s" % (TOPIC_FILE, os.path.relpath(d, ext)))
    print("F1 help links: %d written, %d unchanged, %d removed." % (written, same, removed))


def sidebar_html(current, topics, is_home):
    """current = topic folder name, or None on the home page."""
    home_href = "help.html" if is_home else "../../help.html"
    topic_href = (TOPICS_DIR + "/%s/help.html") if is_home else "../%s/help.html"
    out = ['<nav class="side" aria-label="Help topics">',
           '<input id="topic-filter" type="search" placeholder="Filter topics" aria-label="Filter topics">',
           '<ul id="topics-home">',
           '<li><a class="home-link" href="%s"%s>Home</a></li>' % (
               home_href, ' aria-current="page"' if is_home else ""),
           '</ul><ul id="topics">']
    for folder, title in topics:
        cur = ' aria-current="page"' if folder == current else ""
        out.append('<li><a href="%s"%s>%s</a></li>' % (topic_href % folder, cur, html.escape(title)))
    out.append('</ul><p class="none" id="no-topics">No matching topics.</p></nav>')
    return "\n".join(out)


def prepare_logo(help_root):
    """Copies murray_logo.png (kept next to this script) into the help folder once,
    so every page can reference the same file instead of embedding it."""
    src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "murray_logo.png")
    dst = os.path.join(help_root, "murray_logo.png")
    if os.path.exists(src):
        if os.path.abspath(src) != os.path.abspath(dst):
            shutil.copyfile(src, dst)
        return True
    return os.path.exists(dst)


def build_page(md_path, topics, logo, is_home, info=None, home_data=None):
    page_dir = os.path.dirname(os.path.abspath(md_path))
    current = None if is_home else os.path.basename(page_dir)
    with open(md_path, encoding="utf-8") as fh:
        md = split_header(fh.read())[1]
    warnings = []
    title, body = md_to_html(md, page_dir, warnings)
    meta_tags = ""
    if info:
        meta_tags = ('<meta name="murray-topic-id" content="%s">\n<meta name="murray-content-hash" content="%s">'
                     % (html.escape(info["id"]), info["hash"]))
    logo_tag = ""
    if logo:
        logo_tag = '<img class="band-logo" src="%smurray_logo.png" alt="Murray Company">' % (
            "" if is_home else "../../")
    if is_home:
        home = "Murray Help"
        body = '<div class="home"><div class="page">%s</div>%s</div>' % (
            body, FEEDBACK.format(email=html.escape(FEEDBACK_EMAIL), cc=html.escape(FEEDBACK_CC)))
        if home_data is not None:
            body += '<script type="application/json" id="topics-data">%s</script>' % home_data
    else:
        home = '<a href="../../help.html">Murray Help</a>'
        body = '<div class="page">%s</div>' % body
    page = TEMPLATE.format(title=html.escape(title), css=CSS, js=JS + JS_SEEN, body=body, home=home, logo=logo_tag,
                           meta=meta_tags,
                           sidebar=sidebar_html(current, topics, is_home),
                           date=datetime.date.today().isoformat())
    with open(os.path.join(page_dir, "help.html"), "w", encoding="utf-8") as fh:
        fh.write(page)
    print("Built   " + ("Home" if is_home else current))
    for w in warnings:
        print("   ! " + w)
    return title


def build_all(folder):
    root = find_help_root(folder)
    if not root:
        print("Could not find %s under %s" % (INDEX_FOLDER, folder))
        return
    topics = collect_topics(root)
    logo = prepare_logo(root)
    infos = gather_info(root, topics)
    ext, claims, html_dirs = scan_topic_files(root, topics)
    for folders in claims.values():
        for f in folders:
            infos[f]["type"] = "Tool"
    home_md = os.path.join(root, "help.md")
    if os.path.isfile(home_md):
        build_page(home_md, topics, logo, True, home_data=home_data_json(infos))
    else:
        print("   ! No home page: %s is missing" % home_md)
    for name, _ in topics:
        build_page(os.path.join(root, TOPICS_DIR, name, "help.md"), topics, logo, False, info=infos[name])
    print("%d topics built." % len(topics))
    write_stubs(root, infos, logo, ext, claims, html_dirs)


# --------------------------------------------------------------------------
# Notion HTML export -> Markdown
# --------------------------------------------------------------------------
VOID = {"img", "br", "hr", "meta", "link", "input"}


class Node(object):
    def __init__(self, tag, attrs=None, parent=None):
        self.tag, self.attrs, self.parent, self.children = tag, dict(attrs or {}), parent, []

    def cls(self):
        return self.attrs.get("class", "") or ""

    def find(self, pred):
        for c in self.children:
            if isinstance(c, Node):
                if pred(c):
                    return c
                r = c.find(pred)
                if r:
                    return r
        return None


class TreeBuilder(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self, convert_charrefs=True)
        self.root = Node("root")
        self.cur = self.root

    def handle_starttag(self, tag, attrs):
        n = Node(tag, attrs, self.cur)
        self.cur.children.append(n)
        if tag not in VOID:
            self.cur = n

    def handle_startendtag(self, tag, attrs):
        self.cur.children.append(Node(tag, attrs, self.cur))

    def handle_endtag(self, tag):
        n = self.cur
        while n is not self.root and n.tag != tag:
            n = n.parent
        if n is not self.root:
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.children.append(data)


def text_of(node):
    if isinstance(node, str):
        return node
    return "".join(text_of(c) for c in node.children)


class Importer(object):
    """Converts one Notion page to Markdown. Shared state: page map for cross links."""

    def __init__(self, html_path, bundle_dir, page_map, titles=None):
        self.src_dir = os.path.dirname(html_path)
        self.bundle = bundle_dir
        self.page_map = page_map  # notion id -> bundle folder name
        self.titles = titles or {}  # "../Folder/help.html" -> page title
        self.count = 0
        self.lines = []
        self.warnings = []

    # ---- links and files -------------------------------------------------
    def page_link(self, href):
        m = re.search(r"([0-9a-f]{32})(?:\.html)?(?:[?#].*)?$", unquote(href))
        if m and m.group(1) in self.page_map:
            return "../%s/help.html" % self.page_map[m.group(1)]
        return None

    def copy_file(self, rel, sub):
        path = os.path.join(self.src_dir, unquote(rel))
        if not os.path.isfile(path):
            return None
        base = re.sub(r"[^A-Za-z0-9._-]+", "-", os.path.basename(path)).strip("-")
        os.makedirs(os.path.join(self.bundle, sub), exist_ok=True)
        shutil.copyfile(path, os.path.join(self.bundle, sub, base))
        return "%s/%s" % (sub, base)

    def resolve(self, href, label):
        """Return a Markdown link target, or None to render the label as bold text."""
        if not href or href.startswith("data:"):
            return None
        if "app.notion.comundefined" in href:
            self.warnings.append("Dropped a broken Notion embed link")
            return ""
        if href.startswith(("http://", "https://")):
            if re.search(r"notion\.(so|site|com)", href):
                t = self.page_link(href)
                if t:
                    return t
                self.warnings.append("Notion link to a page not in this export: " + href)
                return None
            return href
        if href.startswith("mailto:"):
            return href
        if href.lower().split("#")[0].endswith(".html"):
            t = self.page_link(href)
            if t:
                return t
            self.warnings.append("Link to page not in export, kept as text: " + label)
            return None
        t = self.copy_file(href, "files")
        if t:
            return t
        self.warnings.append("Linked file not found in export: " + unquote(href))
        return None

    # ---- inline ------------------------------------------------------------
    def inline(self, node, keep_space=False):
        parts = []
        for c in node.children:
            if isinstance(c, str):
                parts.append(re.sub(r"\s+", " ", c))
                continue
            if c.tag in ("ul", "ol", "figure", "div", "p", "h1", "h2", "h3"):
                continue
            inner = self.inline(c, keep_space=True)
            core = inner.strip()
            lead = " " if inner[:1].isspace() else ""
            trail = " " if inner[-1:].isspace() else ""
            if c.tag in ("strong", "b", "mark"):
                parts.append(lead + "**" + core + "**" + trail if core else inner)
            elif c.tag in ("em", "i"):
                parts.append(lead + "*" + core + "*" + trail if core else inner)
            elif c.tag in ("del", "s"):
                parts.append(lead + "~~" + core + "~~" + trail if core else inner)
            elif c.tag == "code":
                parts.append("`%s`" % text_of(c))
            elif c.tag == "a":
                target = self.resolve(c.attrs.get("href", ""), core)
                if target is None:
                    parts.append(lead + "**" + core + "**" + trail if core else "")
                elif target == "":
                    parts.append("")
                else:
                    if target.startswith("../") and (not core or core.startswith("http")):
                        core = self.titles.get(target, core)
                    if core == target and target.startswith(("http://", "https://")):
                        parts.append(lead + target + trail)
                    else:
                        parts.append(lead + "[%s](%s)" % (core or target, target) + trail)
            elif c.tag == "br":
                parts.append("\x01")
            else:
                parts.append(inner)
        s = re.sub(r"[ \t]+", " ", "".join(parts))
        s = s.replace("****", "").replace("** **", " ")
        return s if keep_space else s.strip()

    # ---- blocks ----------------------------------------------------------
    def emit(self, ind, text, note=False):
        """Write a text line; <br> splits become hard line breaks."""
        pieces = [p.strip() for p in text.split("\x01")]
        pieces = [p for p in pieces if p] or []
        for i, p in enumerate(pieces):
            hard = "\\" if (i < len(pieces) - 1 and not note) else ""
            self.lines.append(" " * ind + ("> " if note else "") + p + hard)

    def image(self, fig, ind):
        img = fig.find(lambda n: n.tag == "img")
        if not img:
            return
        src = img.attrs.get("src", "")
        self.count += 1
        os.makedirs(os.path.join(self.bundle, "images"), exist_ok=True)
        if src.startswith("data:"):
            m = re.match(r"data:image/(\w+);base64,(.*)", src, re.S)
            if not m:
                self.warnings.append("Unreadable embedded image skipped")
                return
            ext = "." + m.group(1).replace("jpeg", "jpg")
            name = "img-%02d%s" % (self.count, ext)
            with open(os.path.join(self.bundle, "images", name), "wb") as fh:
                fh.write(base64.b64decode(m.group(2)))
        elif src.startswith("http"):
            self.warnings.append("Online image left as link (not available offline): " + src)
            self.lines.append(" " * ind + "![](%s)" % src)
            return
        else:
            path = os.path.join(self.src_dir, unquote(src))
            if not os.path.isfile(path):
                self.warnings.append("Image not found in export: " + unquote(src))
                return
            ext = os.path.splitext(path)[1].lower() or ".png"
            name = "img-%02d%s" % (self.count, ext)
            shutil.copyfile(path, os.path.join(self.bundle, "images", name))
        cap = fig.find(lambda n: n.tag == "figcaption")
        alt = text_of(cap).strip().replace("]", ")") if cap else ""
        self.lines.append(" " * ind + "![%s](images/%s)" % (alt, name))
        if alt:
            self.lines.append(" " * ind + "*%s*" % alt)

    def blocks(self, node, ind, note=False):
        for c in node.children:
            if isinstance(c, str):
                continue
            k = c.cls()
            colored = re.search(r"block-color-(?!default)", k) is not None
            is_note = note or colored
            if c.tag in ("ul", "ol"):
                n = int(c.attrs.get("start", "1") or 1)
                for li in [x for x in c.children if isinstance(x, Node) and x.tag == "li"]:
                    text = self.inline(li)
                    if is_note:
                        if text:
                            self.emit(ind, text, note=True)
                        self.blocks(li, ind)
                    else:
                        marker = "%d." % n if c.tag == "ol" else "-"
                        first, _, rest = text.partition("\x01")
                        self.lines.append(" " * ind + marker + " " + first.strip())
                        if rest.strip():
                            self.emit(ind + 4, rest)
                        self.blocks(li, ind + 4)
                        n += 1
            elif c.tag == "figure" and "image" in k:
                self.image(c, ind)
            elif c.tag == "figure" and "link-to-page" in k:
                a = c.find(lambda x: x.tag == "a")
                if a:
                    label = text_of(a).strip()
                    t = self.resolve(a.attrs.get("href", ""), label)
                    self.lines.append(" " * ind + ("- [%s](%s)" % (label, t) if t else "- **%s**" % label))
            elif c.tag == "figure":
                a = c.find(lambda x: x.tag == "a")
                if a:
                    href = a.attrs.get("data-notion-file") or a.attrs.get("href", "")
                    label = a.attrs.get("data-notion-file-alias") or text_of(a).strip()
                    t = self.resolve(href, label)
                    if t:
                        src_note = is_note or "block-color-" in (a.parent.cls() if a.parent else "")
                        self.emit(ind, "[%s](%s)" % (label or "Attachment", t), note=src_note)
            elif c.tag in ("h1", "h2", "h3"):
                lvl = "###" if c.tag == "h3" else "##"
                self.lines.append("")
                self.lines.append(lvl + " " + self.inline(c).replace("\x01", " "))
                self.lines.append("")
            elif c.tag == "p":
                t = self.inline(c)
                if t:
                    self.emit(ind, t, note=is_note)
            elif c.tag == "hr":
                self.lines.append("")
                self.lines.append("---")
                self.lines.append("")
            elif c.tag in ("div", "details", "blockquote"):
                self.blocks(c, ind, note=is_note)


def safe_unpack(zip_path, dest):
    with zipfile.ZipFile(zip_path) as z:
        real = os.path.realpath(dest)
        for m in z.infolist():
            if os.path.realpath(os.path.join(dest, m.filename)).startswith(real):
                z.extract(m, dest)
    for root, _, files in os.walk(dest):
        for f in files:
            if f.lower().endswith(".zip"):
                p = os.path.join(root, f)
                sub = os.path.join(root, f[:-4])
                os.makedirs(sub, exist_ok=True)
                safe_unpack(p, sub)
                os.remove(p)


def slug(title):
    words = re.findall(r"[A-Za-z0-9]+", title)
    return "".join(w[:1].upper() + w[1:] for w in words) or "HelpPage"


def import_notion(zip_path, target):
    root = find_help_root(target) or os.path.abspath(target)
    panel_dir = os.path.join(root, TOPICS_DIR)
    work = tempfile.mkdtemp()
    try:
        safe_unpack(zip_path, work)
        os.makedirs(panel_dir, exist_ok=True)
        pages = sorted(os.path.join(r, f) for r, _, fs in os.walk(work)
                       for f in fs if f.lower().endswith(".html"))
        # Pass 1: titles, bundle names, notion ids
        info, page_map, used, titles = [], {}, set(), {}
        for p in pages:
            tb = TreeBuilder()
            with open(p, encoding="utf-8") as fh:
                tb.feed(fh.read())
            h1 = tb.root.find(lambda n: n.tag == "h1" and "page-title" in n.cls())
            title = re.sub(r"\s+", " ", text_of(h1)).strip() if h1 else os.path.splitext(os.path.basename(p))[0]
            name = slug(title)
            n, i = name, 2
            while n in used:
                n, i = "%s%d" % (name, i), i + 1
            used.add(n)
            folder = n
            m = re.search(r"([0-9a-f]{32})\.html$", p)
            if m:
                page_map[m.group(1)] = folder
            titles["../%s/help.html" % folder] = title
            info.append((p, tb, title, folder))
        # Pass 2: convert, then build once every page exists (cross links)
        built = []
        for p, tb, title, folder in info:
            body = tb.root.find(lambda n: n.tag == "div" and "page-body" in n.cls())
            bundle = os.path.join(panel_dir, folder)
            md_path = os.path.join(bundle, "help.md")
            if os.path.exists(md_path):
                print("Skipped %s (help.md already exists)" % folder)
                continue
            os.makedirs(bundle, exist_ok=True)
            imp = Importer(p, bundle, page_map, titles)
            if body:
                imp.blocks(body, 0)
            spaced = []
            for ln in imp.lines:
                if (ln and not ln[0].isspace() and spaced and spaced[-1] != ""
                        and not spaced[-1].endswith("\\")):
                    spaced.append("")
                spaced.append(ln)
            md = "# %s\n\n" % title + "\n".join(spaced).strip() + "\n"
            md = re.sub(r"\n{3,}", "\n\n", md)
            with open(md_path, "w", encoding="utf-8") as fh:
                fh.write(md)
            print("Imported %s -> %s (%d images)" % (title, folder, imp.count))
            for w in imp.warnings:
                print("   ! " + w)
            built.append(md_path)
        print("")
        build_all(root)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "build":
        build_all(sys.argv[2])
    elif len(sys.argv) == 4 and sys.argv[1] == "import":
        import_notion(sys.argv[2], sys.argv[3])
    else:
        # Run from the pyRevit button: build the help site found next to this button
        build_all(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
