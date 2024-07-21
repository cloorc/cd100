#!/usr/bin/env python3
# --encoding: utf-8--
import io
import os
import sys
import time
from typing import List
import urllib3
import bs4
from docx import Document
from htmldocx import HtmlToDocx
from cssutils import parseStyle


server = 'http://www.dagongjibx.com'
index_template = server + '/Article-index-p-%d.html'
eol = {
    '？': True, '。': True, '！': True,
    '?': True, '.': True, '!': True
}


wait = 12


def check_and_decompose(i: bs4.Tag | None, l: List[bs4.Tag]) -> List[bs4.Tag]:
    if i.name == 'img' or i.name == 'br' or len(i.get_text().strip()) <= 0:
        i.decompose()
        if i.parent is not None:
            # should check again
            l.append(i.parent)
    return l


def fetch_and_trim(url: str) -> str:
    retries = 3
    buf = io.BytesIO()
    try:
        res = urllib3.request('GET', url, preload_content=False)
        if 200 <= res.status < 300:
            for chunk in res.stream(5*1024*1024):
                buf.write(chunk)
        else:
            print('Unable to fetch: %s, status=%d, message=%s' % (url, res.status, res.reason))
    except Exception as e:
        if retries <= 0:
            return ''
        else:
            sleep = wait / retries
            print('Exception: %s, will sleep %d seconds and retry ... ' % (str(e), sleep))
            time.sleep(sleep)
            retries -= 1

    res.release_conn()

    bs = bs4.BeautifulSoup(buf.getvalue(), 'html.parser')
    body = bs.select_one('div.news-nr-box')

    text = io.StringIO()
    text.writelines(['\r\n' + str(body.select_one('h1')) + '\r\n'])
    for p in body.find_all_next('p'):
        empties = []
        for i in p.find_all_next():
            empties = check_and_decompose(i, empties)
        while len(empties) > 0:
            i = empties.pop()
            empties = check_and_decompose(i, empties)
        if len(p.get_text().strip()) > 0:
            css = parseStyle(p.get('style'))
            css['box-sizing'] = 'content-box'
            del css['line-height']
            del css['margin-bottom']
            del css['text-align']
            p['style'] = css.cssText.replace('\n', '')
            text.writelines([str(p) + '\n'])
    return text.getvalue()


def fetch_urls() -> List[str]:
    articles = []
    i = 0
    while True:
        i += 1
        print('Fetching page %02d ...' % i)
        count = len(articles)
        buf = io.BytesIO()
        res = urllib3.request('GET', index_template % i, preload_content=False)
        if 200 <= res.status < 300:
            for chunk in res.stream(5*1024*1024):
                buf.write(chunk)
        res.release_conn()
        bs = bs4.BeautifulSoup(buf.getvalue(), 'html.parser')
        for a in bs.select('a.Themetxthover'):
            articles.append(server + a.get('href'))
        if len(articles) - count < 12:
            break
    return articles


if __name__ == '__main__':
    html = '100 chinese doctors.html'
    docx = '100 chinese doctors.docx'

    if not os.path.exists(html) or os.path.getsize(html) <= 0:
        print('Starting to fetch articles ... ')
        urls = fetch_urls()
        with open(html, 'w+', encoding='utf-8') as f:
            for i, url in enumerate(urls):
                print('[%03d] Parsing article : %s ...' % (i, url))
                body = fetch_and_trim(url)
                f.write(body)
                f.write('\n')
    else:
        print('Skip fetching from remote.')

    document = Document()
    parser = HtmlToDocx()
    with open(html, 'r', encoding='utf-8') as f:
        while True:
            try:
                l = f.readline()
                if not l:
                    break
                parser.add_html_to_document(html=f.readline(), document=document)
            except Exception as e:
                print(e)
                break
    document.save(docx)
    print('Fetching acomplished!')
