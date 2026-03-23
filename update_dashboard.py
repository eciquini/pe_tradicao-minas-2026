import urllib.request, urllib.parse, json, os, base64

CLIENT_ID     = '066db1f7-cf8e-4513-b8d0-c6b9df79d1d7'
REFRESH_TOKEN = os.environ.get('MS_REFRESH_TOKEN', '')
FILE_ID       = '6D702C26A5C28103!s5f49050597ab401fa245a35b359f4cc4'
HTML_TEMPLATE = 'pe_tradicao_minas_2026_v26 (2).html'
HTML_OUTPUT   = 'index.html'

if not REFRESH_TOKEN:
    raise Exception('MS_REFRESH_TOKEN nao definido')

def get_token():
    data = urllib.parse.urlencode({
        'client_id': CLIENT_ID,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.Read Files.ReadWrite offline_access User.Read'
    }).encode()
    req = urllib.request.Request(
        'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
        data=data, method='POST',
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )
    with urllib.request.urlopen(req) as r:
        resp = json.loads(r.read())
        if 'refresh_token' in resp:
            with open('new_refresh_token.txt', 'w') as f:
                f.write(resp['refresh_token'])
        return resp['access_token']

def download_excel(token):
    # Pegar URL de download direto via metadados
    url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + FILE_ID
    req = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + token})
    with urllib.request.urlopen(req) as r:
        item = json.loads(r.read())
    
    download_url = item.get('@microsoft.graph.downloadUrl')
    if not download_url:
        raise Exception('downloadUrl nao encontrado nos metadados')
    
    print('Download URL obtida, baixando...')
    req2 = urllib.request.Request(download_url)
    with urllib.request.urlopen(req2) as r:
        return r.read()

def inject_into_html(xlsx_bytes):
    b64 = base64.b64encode(xlsx_bytes).decode()
    with open(HTML_TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()
    auto_script = (
        '\n<script>\n(function(){\n'
        '  var b64="' + b64 + '";\n'
        '  var bin=atob(b64);\n'
        '  var arr=new Uint8Array(bin.length);\n'
        '  for(var i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);\n'
        '  try{\n'
        '    var wb=XLSX.read(arr.buffer,{type:"array",cellFormula:true,cellNF:false,cellHTML:false,cellComment:true});\n'
        '    try{updApplyConfig(wb);}catch(e){}\n'
        '    var result=updParse(wb);\n'
        '    if(result&&result.iraw&&result.iraw.length>0){_updParsed=result;updConfirm();}\n'
        '  }catch(e){console.warn("Auto-update falhou:",e);}\n'
        '})();\n</script>'
    )
    html = html.replace('</body>', auto_script + '\n</body>')
    with open(HTML_OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print('Dashboard gerado: ' + HTML_OUTPUT)

print('Obtendo token...')
token = get_token()
print('Baixando Excel...')
xlsx = download_excel(token)
print('Excel: ' + str(len(xlsx)) + ' bytes')
print('Gerando dashboard...')
inject_into_html(xlsx)
print('Concluido!')
