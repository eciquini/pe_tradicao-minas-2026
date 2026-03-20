import urllib.request, urllib.parse, json, os, io, base64

CLIENT_ID     = '066db1f7-cf8e-4513-b8d0-c6b9df79d1d7'
CLIENT_SECRET = 'W0N8Q~mvr5HVSzAFpS34hTaf2luz3pySsGKDTcdX'
FILE_ID       = 'E9EA55FE-F4A4-4608-80BC-F25D845DFE70'
REFRESH_TOKEN = os.environ.get('MS_REFRESH_TOKEN', '')
HTML_TEMPLATE = 'pe_tradicao_minas_2026_v26 (2).html'
HTML_OUTPUT   = 'index.html'

if not REFRESH_TOKEN:
    raise Exception('MS_REFRESH_TOKEN nao definido')

def get_token():
    data = urllib.parse.urlencode({
        'client_id': CLIENT_ID, 'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token',
        'scope': 'Files.Read offline_access User.Read'
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
    url = 'https://graph.microsoft.com/v1.0/me/drive/items/' + FILE_ID + '/content'
    req = urllib.request.Request(url, headers={'Authorization': 'Bearer ' + token})
    with urllib.request.urlopen(req) as r:
        return r.read()

def inject_into_html(xlsx_bytes):
    b64 = base64.b64encode(xlsx_bytes).decode()
    with open(HTML_TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()
    auto_script = '''
<script>
(function(){
  var b64="''' + b64 + '''";
  var bin=atob(b64);
  var arr=new Uint8Array(bin.length);
  for(var i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
  try{
    var wb=XLSX.read(arr.buffer,{type:'array',cellFormula:true,cellNF:false,cellHTML:false,cellComment:true});
    try{updApplyConfig(wb);}catch(e){}
    var result=updParse(wb);
    if(result&&result.iraw&&result.iraw.length>0){
      _updParsed=result;
      updConfirm();
    }
  }catch(e){console.warn('Auto-update falhou:',e);}
})();
</script>'''
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
