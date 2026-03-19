param(
  [string]$OutputHtml = '',
  [string]$OutputJson = ''
)

$ErrorActionPreference = 'Stop'

if (-not $OutputHtml) {
  $OutputHtml = Join-Path $PSScriptRoot '..\nuskin-colombia-productos.html'
}

if (-not $OutputJson) {
  $OutputJson = Join-Path $PSScriptRoot '..\nuskin-colombia-productos.json'
}

$sourceUrl = 'https://www.nuskin.com/content/nuskin/es_CO/products/shop/view_all.html'
$graphqlUrl = 'https://apis.nuskin.com/product/graphql'

function Normalize-Text {
  param([AllowNull()][string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ''
  }

  $text = [System.Net.WebUtility]::HtmlDecode($Value)
  $text = $text -replace '<[^>]+>', ' '
  $text = $text -replace '\[(.*?)\]\((.*?)\)', '$1'
  $text = $text -replace '(?m)^[\*\-\u2022]\s*', ''
  $text = $text -replace '(?m)^\d+\.\s*', ''
  $text = $text -replace '[`#_~]+', ' '
  $text = $text -replace '\s+', ' '

  return $text.Trim()
}

function Shorten-Text {
  param(
    [AllowNull()][string]$Value,
    [int]$MaxLength = 180
  )

  $text = Normalize-Text $Value
  if (-not $text) {
    return ''
  }

  $sentence = [regex]::Match($text, "^.{1,$MaxLength}(?:[.!?](?:\s|$)|$)").Value.Trim()
  if ($sentence -and $sentence.Length -le $MaxLength) {
    return $sentence
  }

  if ($text.Length -le $MaxLength) {
    return $text
  }

  $slice = $text.Substring(0, $MaxLength)
  $lastSpace = $slice.LastIndexOf(' ')
  if ($lastSpace -gt 80) {
    $slice = $slice.Substring(0, $lastSpace)
  }

  return ($slice.TrimEnd() + '...')
}

function Test-UsefulText {
  param([AllowNull()][string]$Value)

  $text = Normalize-Text $Value
  if (-not $text) {
    return $false
  }

  if ($text.Length -lt 20) {
    return $false
  }

  if ($text -match '^(short description|description)$') {
    return $false
  }

  return $true
}

function Get-FallbackDescription {
  param([pscustomobject]$Product)

  $categoryTail = (($Product.Category -split '\s*/\s*') | Select-Object -Last 1)

  switch -Regex ($Product.Category) {
    'MATERIALES DE APOYO' { return 'Material de apoyo oficial para acompanar presentaciones, promocion o trabajo comercial.' }
    'AP-24' { return 'Producto de higiene oral de la linea AP-24 para el cuidado diario.' }
    'Nutricentials' { return 'Producto de cuidado facial de la linea Nutricentials para complementar la rutina diaria.' }
    'Epoch' { return 'Producto de cuidado personal de la linea Epoch pensado para bienestar y uso cotidiano.' }
    'ageLOC' { return 'Producto de la linea ageLOC enfocado en cuidado personal y apariencia saludable de la piel.' }
    'Tru Face' { return 'Producto de la linea Tru Face enfocado en el cuidado del rostro.' }
    'Galvanic Spa' { return 'Producto o accesorio de la linea Galvanic Spa para complementar la rutina de cuidado.' }
    'Humectantes' { return 'Producto humectante pensado para apoyar la hidratacion diaria de la piel.' }
    'Mascarillas y Exfoliantes' { return 'Producto para complementar la rutina facial con limpieza o renovacion.' }
    'Cuidado Corporal' { return 'Producto de cuidado corporal para apoyar la rutina diaria.' }
    'Pharmanex|TR90|Mas Productos' { return 'Producto de la linea Pharmanex orientado a nutricion, bienestar o apoyo de estilo de vida.' }
    default { return "Producto oficial de Nu Skin Colombia dentro de la categoria $categoryTail." }
  }
}

function Get-ChunkedArray {
  param(
    [object[]]$Items,
    [int]$ChunkSize = 10
  )

  $chunks = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $Items.Count; $i += $ChunkSize) {
    $end = [Math]::Min($i + $ChunkSize - 1, $Items.Count - 1)
    $chunks.Add(@($Items[$i..$end])) | Out-Null
  }

  return $chunks
}

function Invoke-GraphQuery {
  param([string]$Query)

  $body = @{ query = $Query } | ConvertTo-Json -Compress
  return (Invoke-WebRequest -Method Post -Uri $graphqlUrl -ContentType 'application/json' -Body $body).Content | ConvertFrom-Json
}

$response = Invoke-WebRequest -Uri $sourceUrl
$content = $response.Content

$catalogStart = [regex]::Match($content, '<h1[^>]*>\s*Nu Skin Products\s*</h1>', 'IgnoreCase,Singleline')
if (-not $catalogStart.Success) {
  throw 'No se pudo encontrar el inicio del catalogo oficial de Nu Skin Colombia.'
}

$content = $content.Substring($catalogStart.Index)
$pattern = [regex]'(?is)<h1[^>]*>(?<h1>.*?)</h1>|<h2[^>]*>(?<h2>.*?)</h2>|<a href="(?<href>/content/nuskin/es_CO/products/shop[^"]+)" title="(?<title>[^"]*)"[^>]*>(?<name>.*?)</a>'

$currentSection = ''
$currentCategory = ''
$items = New-Object System.Collections.Generic.List[object]

foreach ($match in $pattern.Matches($content)) {
  if ($match.Groups['h1'].Success) {
    $currentSection = Normalize-Text $match.Groups['h1'].Value
    continue
  }

  if ($match.Groups['h2'].Success) {
    $currentCategory = Normalize-Text $match.Groups['h2'].Value
    continue
  }

  if (-not ($match.Groups['href'].Success -and $currentSection -and $currentCategory)) {
    continue
  }

  $name = Normalize-Text $match.Groups['name'].Value
  if (-not $name) {
    $name = Normalize-Text $match.Groups['title'].Value
  }

  $link = 'https://www.nuskin.com' + $match.Groups['href'].Value
  $productId = ''
  if ($link -match '/([0-9]+)\.html$') {
    $productId = $matches[1]
  }

  $items.Add([pscustomobject]@{
    Name = $name
    Category = "$currentSection / $currentCategory"
    Link = $link
    ProductId = $productId
  }) | Out-Null
}

$products = $items |
  Group-Object Link |
  ForEach-Object {
    $first = $_.Group[0]
    [pscustomobject]@{
      Name = $first.Name
      Category = (($_.Group.Category | Where-Object { $_ } | Select-Object -Unique) -join ' | ')
      Link = $first.Link
      ProductId = $first.ProductId
    }
  } |
  Sort-Object Category, Name

$graphMap = @{}
$numericIds = $products | Where-Object { $_.ProductId } | Select-Object -ExpandProperty ProductId -Unique
$queryTemplate = @'
query {
  productById(id:"__ID__", market:"CO", language:"es") {
    title
    description
    salesText
    productImages { url thumbnail alt }
    features { subtitle features }
    benefits { benefits }
    usage { steps recommendations additionalText markdown }
    variants {
      sku
      title
      description
      productImages { url thumbnail alt }
    }
  }
}
'@

foreach ($numericId in $numericIds) {
  $query = $queryTemplate.Replace('__ID__', $numericId)
  $graphResponse = Invoke-GraphQuery -Query $query
  $graphProduct = $graphResponse.data.productById

  if (-not $graphProduct -or -not $graphProduct.title) {
    continue
  }

  $featureText = ''
  if ($graphProduct.features) {
    $featureText = Normalize-Text ((@($graphProduct.features[0].subtitle) + @($graphProduct.features[0].features)) -join ' ')
  }

  $benefitsText = ''
  if ($graphProduct.benefits) {
    $benefitsText = Normalize-Text $graphProduct.benefits.benefits
  }

  $usageText = ''
  if ($graphProduct.usage) {
    $usageText = Normalize-Text ((@($graphProduct.usage.steps) + @($graphProduct.usage.recommendations) + @($graphProduct.usage.additionalText) + @($graphProduct.usage.markdown)) -join ' ')
  }

  $productLevelImage = @($graphProduct.productImages | Where-Object { $_.url } | Select-Object -First 1)[0]
  $variants = @($graphProduct.variants)
  if (-not $variants.Count) {
    $variants = @([pscustomobject]@{
      sku = $numericId
      title = $graphProduct.title
      description = ''
      productImages = @()
    })
  }

  foreach ($variant in $variants) {
    if (-not $variant.sku) {
      continue
    }

    $variantImage = @($variant.productImages | Where-Object { $_.url } | Select-Object -First 1)[0]
    $imageUrl = ''
    $thumbUrl = ''

    if ($variantImage) {
      $imageUrl = $variantImage.url
      $thumbUrl = $variantImage.thumbnail
    } elseif ($productLevelImage) {
      $imageUrl = $productLevelImage.url
      $thumbUrl = $productLevelImage.thumbnail
    }

    $graphMap[$variant.sku] = [pscustomobject]@{
      Title = Normalize-Text $graphProduct.title
      Description = Normalize-Text $graphProduct.description
      VariantDescription = Normalize-Text $variant.description
      SalesText = Normalize-Text $graphProduct.salesText
      FeatureText = $featureText
      BenefitsText = $benefitsText
      UsageText = $usageText
      ImageUrl = $imageUrl
      ThumbnailUrl = $thumbUrl
    }
  }
}

$enrichedProducts = for ($i = 0; $i -lt $products.Count; $i++) {
  $product = $products[$i]
  $graph = $null
  if ($product.ProductId -and $graphMap.ContainsKey($product.ProductId)) {
    $graph = $graphMap[$product.ProductId]
  }

  $briefDescription = ''
  if ($graph) {
    $candidate = @(
      $graph.Description,
      $graph.VariantDescription,
      $graph.SalesText,
      $graph.FeatureText,
      $graph.BenefitsText,
      $graph.UsageText
    ) | Where-Object { Test-UsefulText $_ } | Select-Object -First 1

    if ($candidate) {
      $briefDescription = Shorten-Text $candidate
    }
  }

  if ((-not $briefDescription) -or $briefDescription.Length -lt 20) {
    $briefDescription = Get-FallbackDescription -Product $product
  }

  [pscustomobject]@{
    Number = $i + 1
    Name = $product.Name
    Category = $product.Category
    Description = $briefDescription
    Link = $product.Link
    ProductId = $product.ProductId
    ImageUrl = if ($graph) { $graph.ImageUrl } else { '' }
    ThumbnailUrl = if ($graph) { $graph.ThumbnailUrl } else { '' }
    HasOfficialImage = [bool]($graph -and $graph.ImageUrl)
  }
}

$updatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
$payload = [pscustomobject]@{
  sourceCatalog = $sourceUrl
  sourceGraphql = $graphqlUrl
  updatedAt = $updatedAt
  count = $enrichedProducts.Count
  officialImageCount = @($enrichedProducts | Where-Object { $_.HasOfficialImage }).Count
  products = $enrichedProducts
}

$productsJson = $enrichedProducts | ConvertTo-Json -Depth 6 -Compress

$template = @'
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Indice oficial de productos Nu Skin Colombia</title>
  <meta name="description" content="Indice de productos ofrecidos por Nu Skin Colombia con numero, categoria, descripcion breve, enlace oficial y boton para copiar.">
  <style>
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #fafbfc; color: #23272f; margin: 0; padding: 0; }
  .container { max-width: 1380px; margin: 40px auto; background: #fff; box-shadow: 0 8px 32px rgba(0,0,0,0.08); border-radius: 10px; padding: 32px 40px; }
  h1 { color: #1565c0; margin-top: 0; }
  h2 { color: #2e7d32; margin-top: 32px; }
  h3 { color: #7b1fa2; margin-top: 24px; }
  .note { background: #e3f2fd; border-left: 4px solid #1976d2; padding: 16px 20px; margin: 18px 0; border-radius: 7px; }
  .note ul { list-style: disc inside; margin: 0; padding-left: 1.2em; }
  .note li { margin-bottom: 3px; line-height: 1.35; }
  .note p { margin: 0 0 6px 0; line-height: 1.5; }
  .footer { text-align: right; color: #888; font-size: 1em; margin-top: 40px; }
  .tag { background: #e3f2fd; color: #1565c0; border-radius: 4px; padding: 3px 8px; margin-left: 6px; font-size: 0.95em; }
  dt { color: #2e7d32; font-weight: bold; margin-top: 10px; }
  dd { margin-bottom: 10px; }
  .params table { border-collapse: collapse; width: 100%; margin: 10px 0; }
  .params th, .params td { border: 1px solid #ccc; padding: 8px; vertical-align: top; }
  .important { background: #fffde7; border-left: 4px solid #fbc02d; padding: 12px 18px; margin: 18px 0; border-radius: 5px; }
  .sideEffects ul, .solid ul { margin: 0; padding-left: 20px; }
  @media (max-width: 600px) { .container { padding: 16px 8px; } }
  pre { background: #f6f8fa; padding: 14px; border-radius: 8px; overflow: auto; font-size: 0.92em; }
  code { font-family: Consolas, 'Courier New', monospace; }
  .row { display:flex; gap:20px; flex-wrap:wrap; }
  .col { flex:1 1 420px; }
  .endpoint { border: 1px solid #e0e0e0; padding: 12px 14px; border-radius: 8px; background: #fff; box-shadow: 0 3px 8px rgba(0,0,0,0.04); }
  .json { white-space: pre; font-family: Consolas, 'Courier New', monospace; font-size: 0.95em; }
  .success { color: #2e7d32; font-weight: 600; }
  .error { color: #c62828; font-weight: 600; }
  .toolbar { display:flex; gap:12px; flex-wrap:wrap; margin: 22px 0 14px; }
  .toolbar input, .toolbar select { flex:1 1 280px; border: 1px solid #cfd8dc; border-radius: 8px; padding: 10px 12px; font: inherit; }
  .params { overflow-x: auto; }
  .name { font-weight: 600; color: #23272f; }
  .small { font-size: 0.92em; color: #666; }
  .copy-btn, .img-btn { border: 0; border-radius: 8px; padding: 8px 10px; font: inherit; cursor: pointer; }
  .copy-btn { background: #1565c0; color: #fff; font-weight: 600; }
  .copy-btn:hover { background: #0d47a1; }
  .img-btn { background: #fff3e0; color: #ef6c00; font-size: 1.1rem; }
  .img-btn:hover { background: #ffe0b2; }
  .img-btn.empty { background: #eceff1; color: #607d8b; }
  .link { color: #1565c0; word-break: break-word; }
  .url { display:block; margin-top: 4px; font-size: 0.88em; color: #607d8b; word-break: break-word; }
  .count-box { display:flex; gap:12px; flex-wrap:wrap; margin-top: 14px; }
  .count-box .endpoint { flex: 1 1 180px; }
  .badge { display:inline-block; border-radius: 999px; padding: 3px 10px; background:#eef7ee; color:#2e7d32; font-size: 0.9em; font-weight: 600; }
  .download-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 18px; margin: 18px 0 22px; }
  .download-card { border: 1px solid #e0e0e0; border-radius: 10px; background: #fff; box-shadow: 0 3px 8px rgba(0,0,0,0.04); padding: 18px; }
  .download-card h3 { margin: 0 0 8px 0; color: #1565c0; font-size: 1.08rem; }
  .download-meta { margin-top: 8px; color: #666; font-size: 0.9rem; }
  .download-actions { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 14px; }
  .download-btn { display: inline-block; padding: 10px 14px; border-radius: 8px; text-decoration: none; font-weight: 600; border: 1px solid #d7e3f4; background: #eef4fb; color: #1565c0; }
  .download-btn.primary { background: #1565c0; color: #fff; border-color: #1565c0; }
  .download-btn:hover { filter: brightness(1.02); }
  .toast { position: fixed; right: 18px; bottom: 18px; background: rgba(35,39,47,0.96); color: #fff; padding: 12px 14px; border-radius: 10px; opacity: 0; transform: translateY(10px); pointer-events: none; transition: opacity 160ms ease, transform 160ms ease; z-index: 60; }
  .toast.show { opacity: 1; transform: translateY(0); }
  .modal[hidden] { display: none; }
  .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.72); display: flex; align-items: center; justify-content: center; padding: 24px; z-index: 70; }
  .modal-card { width: min(900px, 100%); background: #fff; border-radius: 14px; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.25); }
  .modal-head { display:flex; align-items:center; justify-content:space-between; gap:12px; padding: 16px 18px; border-bottom: 1px solid #eceff1; }
  .modal-title { margin: 0; font-size: 1.05rem; color: #1565c0; }
  .modal-close { border: 0; background: #f5f5f5; border-radius: 8px; padding: 8px 10px; cursor: pointer; font: inherit; }
  .modal-body { padding: 18px; }
  .modal-img-wrap { background: #f8fafc; border: 1px solid #eceff1; border-radius: 12px; overflow: hidden; min-height: 320px; display:flex; align-items:center; justify-content:center; }
  .modal-img { display:block; width:100%; height:auto; max-height: 70vh; object-fit: contain; background: #fff; }
  .modal-caption { margin: 14px 0 0 0; color: #555; line-height: 1.6; }
  .sr-only { position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); border:0; }
  </style>
</head>
<body>
  <div class="container">
    <h1>Indice oficial de productos Nu Skin Colombia</h1>
    <p>
      Listado generado desde el catalogo oficial de Nu Skin Colombia. Cada fila incluye numeracion,
      categoria, descripcion breve, enlace oficial, boton para copiar y un emoji para ver imagen.
    </p>

    <div class="note">
      <p><strong>Fuentes oficiales consultadas:</strong></p>
      <ul>
        <li><a href="__SOURCE_CATALOG__" target="_blank" rel="noopener">Catalogo Ver Todos los Productos</a></li>
        <li><code>https://apis.nuskin.com/product/graphql</code> para enriquecer descripcion e imagen cuando estuvo disponible.</li>
      </ul>
    </div>

    <div class="count-box">
      <div class="endpoint">
        <div class="small">Productos indexados</div>
        <div class="success" id="product-count">__COUNT__</div>
      </div>
      <div class="endpoint">
        <div class="small">Visibles en pantalla</div>
        <div class="success" id="visible-count">__COUNT__</div>
      </div>
      <div class="endpoint">
        <div class="small">Con imagen oficial</div>
        <div class="success">__IMAGE_COUNT__</div>
      </div>
      <div class="endpoint">
        <div class="small">Generado</div>
        <div class="small">__UPDATED_AT__</div>
      </div>
    </div>

    <div class="important">
      <p><strong>Nota:</strong> cuando Nu Skin no expone una imagen oficial en la fuente consultada, el emoji abre una vista placeholder para no dejar la fila vacia.</p>
    </div>

    <h2>Descargas</h2>
    <div class="note">
      <p><strong>PDF disponibles:</strong> lista de precios detallada Colombia 2025 y resumen visual del plan de compensacion LATAM 2025.</p>
    </div>

    <div class="download-grid">
      <article class="download-card">
        <span class="badge">PDF / Colombia / 2025</span>
        <h3>Lista de precios detallada Colombia 2025</h3>
        <p>Documento oficial en PDF con la lista de precios detallada para Colombia.</p>
        <div class="download-meta">Archivo: <code>CO-lista-de-precios-detallada-2025.pdf</code> | 0.50 MB</div>
        <div class="download-actions">
          <a class="download-btn primary" href="PDFs/CO-lista-de-precios-detallada-2025.pdf" target="_blank" rel="noopener">Abrir PDF</a>
          <a class="download-btn" href="PDFs/CO-lista-de-precios-detallada-2025.pdf" download>Descargar</a>
        </div>
      </article>

      <article class="download-card">
        <span class="badge">PDF / LATAM / 2025</span>
        <h3>Sales Performance Plan Overview Infographic LATAM 2025</h3>
        <p>Infografia oficial en PDF con el resumen visual del plan de compensacion para LATAM.</p>
        <div class="download-meta">Archivo: <code>sales-performance-plan-overview-infographic-latam-2025.pdf</code> | 0.59 MB</div>
        <div class="download-actions">
          <a class="download-btn primary" href="PDFs/sales-performance-plan-overview-infographic-latam-2025.pdf" target="_blank" rel="noopener">Abrir PDF</a>
          <a class="download-btn" href="PDFs/sales-performance-plan-overview-infographic-latam-2025.pdf" download>Descargar</a>
        </div>
      </article>
    </div>

    <div class="toolbar">
      <input id="search" type="search" placeholder="Buscar por nombre, categoria, descripcion o enlace">
      <select id="category-filter">
        <option value="">Todas las categorias</option>
      </select>
    </div>

    <div class="params">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Nombre</th>
            <th>Categoria</th>
            <th>Descripcion breve</th>
            <th>Enlace oficial</th>
            <th>Copiar</th>
            <th>Imagen</th>
          </tr>
        </thead>
        <tbody id="rows"></tbody>
      </table>
    </div>

    <div id="empty" class="note" hidden>
      <p>No se encontraron resultados con ese filtro.</p>
    </div>

    <div class="footer">
      Actualizado: __UPDATED_AT__
    </div>
  </div>

  <div id="toast" class="toast" role="status" aria-live="polite"></div>

  <div id="image-modal" class="modal" hidden>
    <div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="modal-title">
      <div class="modal-head">
        <h2 id="modal-title" class="modal-title">Vista del producto</h2>
        <button id="modal-close" class="modal-close" type="button">Cerrar</button>
      </div>
      <div class="modal-body">
        <div class="modal-img-wrap">
          <img id="modal-image" class="modal-img" alt="">
        </div>
        <p id="modal-caption" class="modal-caption"></p>
      </div>
    </div>
  </div>

  <script>
    const products = __PRODUCTS__;
    const rows = document.getElementById('rows');
    const empty = document.getElementById('empty');
    const searchInput = document.getElementById('search');
    const categoryFilter = document.getElementById('category-filter');
    const visibleCount = document.getElementById('visible-count');
    const toast = document.getElementById('toast');
    const modal = document.getElementById('image-modal');
    const modalImage = document.getElementById('modal-image');
    const modalTitle = document.getElementById('modal-title');
    const modalCaption = document.getElementById('modal-caption');
    const modalClose = document.getElementById('modal-close');

    const escapeHtml = (value) => String(value).replace(/[&<>"']/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[char]));

    const categories = [...new Set(products.map((product) => product.Category))].sort((a, b) => a.localeCompare(b));
    categories.forEach((category) => {
      const option = document.createElement('option');
      option.value = category;
      option.textContent = category;
      categoryFilter.appendChild(option);
    });

    function showToast(message) {
      toast.textContent = message;
      toast.classList.add('show');
      clearTimeout(showToast.timer);
      showToast.timer = setTimeout(() => toast.classList.remove('show'), 1800);
    }

    async function copyLink(link) {
      try {
        await navigator.clipboard.writeText(link);
        showToast('Enlace copiado');
      } catch (error) {
        const helper = document.createElement('textarea');
        helper.value = link;
        document.body.appendChild(helper);
        helper.select();
        document.execCommand('copy');
        helper.remove();
        showToast('Enlace copiado');
      }
    }

    function buildPlaceholder(product) {
      const title = product.Name;
      const category = product.Category.split(' / ').slice(-1)[0];
      const svg = `
        <svg xmlns="http://www.w3.org/2000/svg" width="1200" height="900" viewBox="0 0 1200 900">
          <defs>
            <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" stop-color="#eef7ff" />
              <stop offset="100%" stop-color="#d8ecff" />
            </linearGradient>
          </defs>
          <rect width="1200" height="900" fill="url(#g)" />
          <circle cx="200" cy="180" r="170" fill="#ffffff" opacity="0.5" />
          <circle cx="980" cy="730" r="190" fill="#ffffff" opacity="0.35" />
          <rect x="110" y="120" width="980" height="660" rx="36" fill="#ffffff" opacity="0.84" />
          <text x="160" y="240" fill="#1565c0" font-family="Segoe UI, Arial, sans-serif" font-size="28" font-weight="700">NU SKIN COLOMBIA</text>
          <text x="160" y="360" fill="#23272f" font-family="Segoe UI, Arial, sans-serif" font-size="56" font-weight="700">${escapeHtml(title)}</text>
          <text x="160" y="450" fill="#2e7d32" font-family="Segoe UI, Arial, sans-serif" font-size="34">${escapeHtml(category)}</text>
          <text x="160" y="560" fill="#607d8b" font-family="Segoe UI, Arial, sans-serif" font-size="28">Vista placeholder cuando no hay imagen oficial disponible.</text>
        </svg>
      `;
      return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`;
    }

    function openImage(product) {
      const imageUrl = product.ImageUrl || buildPlaceholder(product);
      modalImage.src = imageUrl;
      modalImage.alt = product.Name;
      modalTitle.textContent = product.Name;
      modalCaption.textContent = `${product.Category} | ${product.Description}`;
      modal.hidden = false;
      document.body.style.overflow = 'hidden';
    }

    function closeModal() {
      modal.hidden = true;
      modalImage.src = '';
      document.body.style.overflow = '';
    }

    function render() {
      const search = searchInput.value.trim().toLowerCase();
      const selectedCategory = categoryFilter.value;

      const filtered = products.filter((product) => {
        const haystack = `${product.Name} ${product.Category} ${product.Description} ${product.Link}`.toLowerCase();
        const matchesSearch = !search || haystack.includes(search);
        const matchesCategory = !selectedCategory || product.Category === selectedCategory;
        return matchesSearch && matchesCategory;
      });

      visibleCount.textContent = filtered.length;

      if (!filtered.length) {
        rows.innerHTML = '';
        empty.hidden = false;
        return;
      }

      empty.hidden = true;
      rows.innerHTML = filtered.map((product) => `
        <tr>
          <td>${product.Number}</td>
          <td><div class="name">${escapeHtml(product.Name)}</div></td>
          <td><span class="badge">${escapeHtml(product.Category)}</span></td>
          <td>${escapeHtml(product.Description)}</td>
          <td>
            <a class="link" href="${escapeHtml(product.Link)}" target="_blank" rel="noopener">Abrir enlace oficial</a>
            <span class="url">${escapeHtml(product.Link)}</span>
          </td>
          <td><button class="copy-btn" type="button" data-link="${escapeHtml(product.Link)}">Copiar</button></td>
          <td>
            <button
              class="img-btn ${product.HasOfficialImage ? '' : 'empty'}"
              type="button"
              data-number="${product.Number}"
              title="${product.HasOfficialImage ? 'Ver imagen oficial' : 'Ver vista placeholder'}"
            >&#x1F5BC;</button>
          </td>
        </tr>
      `).join('');
    }

    searchInput.addEventListener('input', render);
    categoryFilter.addEventListener('change', render);

    rows.addEventListener('click', (event) => {
      const copyButton = event.target.closest('.copy-btn');
      if (copyButton) {
        copyLink(copyButton.dataset.link);
        return;
      }

      const imageButton = event.target.closest('.img-btn');
      if (imageButton) {
        const number = Number(imageButton.dataset.number);
        const product = products.find((item) => item.Number === number);
        if (product) {
          openImage(product);
        }
      }
    });

    modalClose.addEventListener('click', closeModal);
    modal.addEventListener('click', (event) => {
      if (event.target === modal) {
        closeModal();
      }
    });
    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && !modal.hidden) {
        closeModal();
      }
    });

    render();
  </script>
</body>
</html>
'@

$html = $template.
  Replace('__SOURCE_CATALOG__', $sourceUrl).
  Replace('__UPDATED_AT__', $updatedAt).
  Replace('__COUNT__', [string]$enrichedProducts.Count).
  Replace('__IMAGE_COUNT__', [string](@($enrichedProducts | Where-Object { $_.HasOfficialImage }).Count)).
  Replace('__PRODUCTS__', $productsJson)

$payload | ConvertTo-Json -Depth 8 | Set-Content -Path $OutputJson -Encoding utf8
$html | Set-Content -Path $OutputHtml -Encoding utf8

Write-Output "Productos: $($enrichedProducts.Count)"
Write-Output "Con imagen oficial: $(@($enrichedProducts | Where-Object { $_.HasOfficialImage }).Count)"
Write-Output "HTML: $((Resolve-Path $OutputHtml).Path)"
Write-Output "JSON: $((Resolve-Path $OutputJson).Path)"
Write-Output "Fuente catalogo: $sourceUrl"
Write-Output "Fuente GraphQL: $graphqlUrl"
