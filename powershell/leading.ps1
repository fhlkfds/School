$apiUrl = "https://inv.nomma.lan/api/v1/hardware"
$apiToken = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIzIiwianRpIjoiZWNhZjc1MGVhM2M4OGNkMzc0ZWJhZjc0NGMxMDZhOThkNzYzM2I0NGZhM2RkMjRjMmIwNzZiOWE4OGNmYWVlNTEzNmExNzM1NjU2ZGM3MjgiLCJpYXQiOjE3NDcyMjk0MjYuNjUxODc2LCJuYmYiOjE3NDcyMjk0MjYuNjUxODgzLCJleHAiOjIyMjA2MTUwMjYuNjQyNjExLCJzdWIiOiIxIiwic2NvcGVzIjpbXX0.Dx2UEMszBqPfcNIQbF5qaPTJYd_qEyMUwYJomYJ7h7_CBupOpHhBrr_Bra-CNEAuzGTrd7d5D1gptor0hyKjIk4TLndkMJjFtUPJwqufPa7Hx9NJN3ygvMu3TZWYKGD6iY_GIZj07ljXZp6Ih4XUkIOr8YcPsybOeqF73ovsZD_1950IwTV1OUSNoSK42NWf5U92wNOgP4G9Hd_sclWat0WWnxEeEWTrl9bcGoBiYipQ2wbYX36KIe8eK7KqHVR_5HYYcaPy6fuqCKY5iHwEW2PJluaNvYhIThAE4E9QrgJ-4smipj_iNxAFaf46ESY7bvV0B6ed_d-iPNmm9E73t_UhuIkxGIfN9TRdt1dCteDJFx3qJvi8y9cfY6E2Pws8p2OKzGRLiAeUPRgRUAbj2EjRNNZU7Z0J5cbb-LHTYRBYpT6aYq7zMwN5ZNTfCSjv7OZrhVUoqyAwu1w0FKxrhJQRHm98nOv6hx6DMhEN0kRCPYUVplH6yd2jhNmhlhmUhnihNxAYRVwmxMhxHFIH_q6v7bpc64zCw85vbkee3ETIzq5KxxLD6E5EqFLDL1vifNLMvaVC-wB1-P46DGoxdVXSGizQdvXYnhq5otLQMj9Ryzv6QqSyEN_kXldh0VrAmCBE9rStLUdwErLU4GhrTkDRsTqf-FhDHumL76w8Tv0"

# Get all assets
$response = Invoke-RestMethod -Uri "$apiUrl?limit=5000" -Headers @{Authorization = "Bearer $apiToken"}

foreach ($asset in $response.rows)
{
  if ($asset.asset_tag -match '^[0-9]{1,3}$')
  {
    $newTag = $asset.asset_tag.PadLeft(4, '0')
    Write-Host "Updating Asset ID $($asset.id) - New Tag: $newTag"
    Invoke-RestMethod -Method PATCH -Uri "$apiUrl/$($asset.id)" -Headers @{Authorization = "Bearer $apiToken"} -Body @{asset_tag = $newTag}
  }
}

