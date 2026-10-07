"""Offline DOM checks of the setup page. No browser navigation or network.
Production script imports are joined in memory, location and fetch are mocked.
This does not test GitHub hosting, Office integration, or a live mailbox.
"""
from pathlib import Path
import base64
import json
import mimetypes
import os
import re
from lxml import etree
from playwright.sync_api import sync_playwright

root=Path(__file__).resolve().parents[1]
site=root/'site'
report=root/'verification'
report.mkdir(exist_ok=True)
base='https://example-user.github.io/message-notes/'
html=(site/'index.html').read_text()
html=re.sub(r'<script[^>]*>.*?</script>','',html,flags=re.S)
html=re.sub(r'<link[^>]*>','',html)
html=html.replace('</head>','<style>'+(site/'styles.css').read_text()+'</style></head>')
icon=base64.b64encode((site/'assets/icon-32.png').read_bytes()).decode()
html=html.replace('./assets/icon-32.png','data:image/png;base64,'+icon)
files={}
for f in site.rglob('*'):
    if f.is_file():
        rel=f.relative_to(site).as_posix()
        files[rel]={'body': f.read_text() if f.suffix!='.png' else '', 'type': mimetypes.guess_type(str(f))[0] or 'application/octet-stream'}
builder=(site/'manifest-builder.js').read_text().replace('export const ','const ').replace('export function ','function ')
setup=(site/'setup.js').read_text()
setup=re.sub(r'^import .*?;\n','',setup,count=1)
setup=setup.replace('new URL(location.href)','new URL(window.__fixtureAddress)')
checks=[]
errors=[]
with sync_playwright() as p:
    browser=p.chromium.launch(headless=True,executable_path=os.environ.get('CHROMIUM_PATH','/usr/bin/chromium'),args=['--no-sandbox'])
    page=browser.new_page(viewport={'width':1060,'height':860})
    page.on('pageerror',lambda error:errors.append(str(error)))
    page.set_content(html)
    page.evaluate('v => { window.__files=v.files; window.__fixtureAddress=v.base; }',{'files':files,'base':base})
    page.evaluate('''() => {
      window.fetch = async value => {
        const url=new URL(String(value));
        const rel=url.pathname.replace(/^\\/message-notes\\//,'');
        const file=window.__files[rel];
        return new Response(file ? file.body : 'Missing', {status:file ? 200 : 404,headers:{'content-type':file ? file.type : 'text/plain'}});
      };
      URL.createObjectURL=blob => { window.__downloadBlob=blob; return 'blob:offline-test'; };
      URL.revokeObjectURL=()=>{};
      HTMLAnchorElement.prototype.click=function(){ window.__downloadName=this.download; };
    }''')
    page.add_script_tag(content='(function(){\n'+builder+'\n'+setup+'\nwindow.__prepareTest=prepare;\n})();')
    page.wait_for_function("!document.getElementById('downloadManifest').disabled")
    assert page.locator('#siteAddress').inner_text()==base
    checks.append('Setup detects the mocked hosted base and enables download after simulated asset checks.')
    page.locator('#downloadManifest').click()
    xml=page.evaluate('async () => await window.__downloadBlob.text()')
    assert page.evaluate('window.__downloadName')=='Message-Notes-GitHub.xml'
    parsed=etree.fromstring(xml.encode())
    ns={'o':'http://schemas.microsoft.com/office/appforoffice/1.1','v0':'http://schemas.microsoft.com/office/mailappversionoverrides','v1':'http://schemas.microsoft.com/office/mailappversionoverrides/1.1'}
    assert parsed.find('o:Id',ns).text=='7c6010d9-3b66-45f7-86ca-296e175b2b86'
    assert parsed.find('o:Version',ns).text=='1.0.2.0'
    assert parsed.find('o:Permissions',ns).text=='ReadItem'
    assert parsed.find('.//v0:VersionOverrides/v1:VersionOverrides',ns) is not None
    assert parsed.find('.//v1:SupportsPinning',ns).text=='true'
    assert not re.search('__BASE_URL__|__ORIGIN__|localhost',xml)
    for value in parsed.xpath('//@DefaultValue'):
        if value.startswith('https://'):
            assert value.startswith(base)
            assert (site/value.removeprefix(base)).is_file()
    # Kept only in verification, never deployed or offered as an installation file.
    (report/'simulated-host-manifest.xml.example').write_text(xml)
    checks.append('Download action creates correctly named XML with the same ID, ReadItem, pinning, and complete hosted asset paths.')
    page.screenshot(path=str(report/'hosted-setup.png'),full_page=True)
    page.set_viewport_size({'width':320,'height':760})
    assert page.evaluate('document.documentElement.scrollWidth<=window.innerWidth')
    checks.append('Setup layout fits 320 pixels without horizontal overflow.')
    page.evaluate("async () => { window.__fixtureAddress='https://example-user.github.io/message-notes/index.html?x=1#step'; await window.__prepareTest(); }")
    assert page.locator('#siteAddress').inner_text()==base
    assert page.locator('#downloadManifest').is_enabled()
    checks.append('Query and fragment on the setup page do not leak into generated resource addresses.')
    page.evaluate("async () => { delete window.__files['assets/icon-80.png']; await window.__prepareTest(); }")
    assert page.locator('#downloadManifest').is_disabled()
    assert 'HTTP 404' in page.locator('#setupStatus').inner_text()
    checks.append('A missing asset disables download and identifies the missing file.')
    page.evaluate("async () => { window.__fixtureAddress='https://localhost:3000/'; await window.__prepareTest(); }")
    assert page.locator('#downloadManifest').is_disabled()
    assert 'localhost' in page.locator('#setupStatus').inner_text()
    checks.append('A localhost setup address is rejected with an actionable error.')
    assert not errors,errors
    checks.append('No browser JavaScript exceptions in the offline DOM tests.')
    browser.close()
result={'passed':len(checks),'failed':0,'checks':checks,'scope':'Offline DOM tests in Chromium about:blank using in-memory production JS, a mocked address, and mocked fetch. Download bytes were inspected, not saved by a real browser download. No external hosting or Outlook integration verified.'}
(report/'browser-checks.json').write_text(json.dumps(result,indent=2))
print(json.dumps(result,indent=2))
