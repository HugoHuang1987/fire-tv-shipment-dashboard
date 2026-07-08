param(
    [int]$TimeoutSeconds = 30
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http
$baseUrl = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_BASE_URL', 'Process')
if (-not $baseUrl) { $baseUrl = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_BASE_URL', 'User') }
$model = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_MODEL', 'Process')
if (-not $model) { $model = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_MODEL', 'User') }
$apiKey = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_API_KEY', 'Process')
if (-not $apiKey) { $apiKey = [Environment]::GetEnvironmentVariable('FIRE_TV_LLM_API_KEY', 'User') }

if (-not $baseUrl -or -not $model -or -not $apiKey) {
    throw 'LLM环境变量不完整：需要FIRE_TV_LLM_BASE_URL、FIRE_TV_LLM_MODEL、FIRE_TV_LLM_API_KEY。'
}

$uri = "$($baseUrl.TrimEnd('/'))/chat/completions"
$body = @{
    model = $model
    messages = @(
        @{ role = 'system'; content = '你是连通性测试助手。只回复JSON：{"status":"ok"}' },
        @{ role = 'user'; content = '请确认接口可用。' }
    )
    stream = $false
} | ConvertTo-Json -Depth 10

$handler = New-Object System.Net.Http.HttpClientHandler
$handler.UseProxy = $false
$client = New-Object System.Net.Http.HttpClient($handler)
$client.Timeout = [TimeSpan]::FromSeconds($TimeoutSeconds)
$client.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $apiKey)
try {
    $content = New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, 'application/json')
    $response = $client.PostAsync($uri, $content).GetAwaiter().GetResult()
    $responseText = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
    if (-not $response.IsSuccessStatusCode) {
        throw "LLM接口返回HTTP $([int]$response.StatusCode)：$responseText"
    }
    $parsed = $responseText | ConvertFrom-Json
    $reply = $parsed.choices[0].message.content
    [pscustomobject]@{
        Connected = $true
        Endpoint = $uri
        Model = $model
        ReplyPresent = [bool]$reply
    }
} finally {
    $client.Dispose()
    $handler.Dispose()
}
