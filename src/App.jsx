import { useState, useRef, useCallback } from 'react'
import './App.css'

const CATEGORIES = {
  pdf: {
    label: 'PDF',
    color: '#e02020',
    icon: 'PDF',
    formats: ['.pdf'],
  },
  word: {
    label: 'Word',
    color: '#2b5797',
    icon: 'W',
    formats: ['.docx', '.doc', '.odt'],
  },
  excel: {
    label: 'Excel',
    color: '#1e6b3c',
    icon: 'X',
    formats: ['.xlsx', '.xls', '.ods', '.csv'],
  },
  powerpoint: {
    label: 'PowerPoint',
    color: '#c43e1c',
    icon: 'P',
    formats: ['.pptx', '.ppt', '.odp'],
  },
  pictures: {
    label: 'Imagen',
    color: '#0077c8',
    icon: 'IMG',
    formats: ['.jpg', '.jpeg', '.png', '.svg'],
  },
}

const MIME_TO_CATEGORY = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'word',
  'application/msword': 'word',
  'application/vnd.oasis.opendocument.text': 'word',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'excel',
  'application/vnd.ms-excel': 'excel',
  'application/vnd.oasis.opendocument.spreadsheet': 'excel',
  'text/csv': 'excel',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'powerpoint',
  'application/vnd.ms-powerpoint': 'powerpoint',
  'application/vnd.oasis.opendocument.presentation': 'powerpoint',
  'application/pdf': 'pdf',
  'image/jpeg': 'pictures',
  'image/png': 'pictures',
  'image/svg+xml': 'pictures',
}

const EXT_TO_CATEGORY = {
  pdf: 'pdf',
  docx: 'word', doc: 'word', odt: 'word',
  xlsx: 'excel', xls: 'excel', ods: 'excel', csv: 'excel',
  pptx: 'powerpoint', ppt: 'powerpoint', odp: 'powerpoint',
  jpg: 'pictures', jpeg: 'pictures', png: 'pictures', svg: 'pictures',
}

function getCategoryForFile(file) {
  if (MIME_TO_CATEGORY[file.type]) return MIME_TO_CATEGORY[file.type]
  const ext = file.name.split('.').pop().toLowerCase()
  return EXT_TO_CATEGORY[ext] || null
}

function getExtForFile(file) {
  return '.' + file.name.split('.').pop().toLowerCase()
}

const ALL_ACCEPTED = '.pdf,.docx,.doc,.odt,.xlsx,.xls,.ods,.csv,.pptx,.ppt,.odp,.jpg,.jpeg,.png,.svg'

function formatBytes(bytes) {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`
}

function FileCard({ file, onRemove, targetFormat, onFormatChange, converting }) {
  const catKey = getCategoryForFile(file)
  const cat = catKey ? CATEGORIES[catKey] : null
  const currentExt = getExtForFile(file)
  const isImage = catKey === 'pictures'
  const [preview, setPreview] = useState(null)

  if (isImage && !preview) {
    const reader = new FileReader()
    reader.onload = (e) => setPreview(e.target.result)
    reader.readAsDataURL(file)
  }

  const convertOptions = cat
    ? cat.formats.filter(f => f !== currentExt)
    : []

  return (
    <div className={`file-card${converting ? ' file-card--converting' : ''}`}>
      <div className="file-card-preview" style={{ background: isImage ? '#111' : (cat ? cat.color + '18' : '#6b728018') }}>
        {isImage && preview
          ? <img src={preview} alt={file.name} className="file-thumb" />
          : <span className="file-icon" style={{ color: cat ? cat.color : '#6b7280' }}>{cat ? cat.icon : '?'}</span>
        }
      </div>
      <div className="file-card-body">
        <p className="file-name" title={file.name}>{file.name}</p>
        <div className="file-meta">
          <span className="file-badge" style={{ background: cat ? cat.color : '#6b7280' }}>
            {currentExt.replace('.', '').toUpperCase()}
          </span>
          <span className="file-type-label">{cat ? cat.label : 'Desconocido'}</span>
        </div>
        <p className="file-size">{formatBytes(file.size)}</p>

        {cat && convertOptions.length > 0 && (
          <div className="convert-row">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none" className="convert-arrow">
              <path d="M3 8h10M9 4l4 4-4 4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
            <div className="convert-options">
              {convertOptions.map(fmt => (
                <button
                  key={fmt}
                  className={`convert-btn${targetFormat === fmt ? ' convert-btn--active' : ''}`}
                  style={{
                    '--btn-color': cat.color,
                    '--btn-bg': cat.color + '14',
                    '--btn-bg-active': cat.color + '28',
                  }}
                  onClick={() => onFormatChange(fmt)}
                  disabled={converting}
                >
                  {fmt.replace('.', '').toUpperCase()}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
      <button className="file-remove" onClick={() => onRemove(file)} title="Eliminar" disabled={converting}>
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
          <path d="M11 3L3 11M3 3l8 8" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/>
        </svg>
      </button>
    </div>
  )
}

export default function App() {
  const [files, setFiles] = useState([])
  const [targetFormats, setTargetFormats] = useState({})
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState('')
  const [converting, setConverting] = useState(false)
  const [convertError, setConvertError] = useState('')
  const inputRef = useRef()

  const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
  const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY

  const addFiles = useCallback((incoming) => {
    setError('')
    const valid = []
    const invalid = []
    for (const f of incoming) {
      if (getCategoryForFile(f)) valid.push(f)
      else invalid.push(f.name)
    }
    if (invalid.length) setError(`Formato no soportado: ${invalid.join(', ')}`)
    if (valid.length) {
      setFiles(prev => {
        const seen = new Set(prev.map(f => f.name + f.size))
        return [...prev, ...valid.filter(f => !seen.has(f.name + f.size))]
      })
    }
  }, [])

  const onDrop = useCallback((e) => {
    e.preventDefault()
    setDragging(false)
    addFiles(Array.from(e.dataTransfer.files))
  }, [addFiles])

  const onDragOver = (e) => { e.preventDefault(); setDragging(true) }
  const onDragLeave = () => setDragging(false)
  const onInputChange = (e) => { addFiles(Array.from(e.target.files)); e.target.value = '' }

  const removeFile = (file) => {
    setFiles(prev => prev.filter(f => f !== file))
    setTargetFormats(prev => {
      const next = { ...prev }
      delete next[file.name + file.size]
      return next
    })
  }

  const clearAll = () => { setFiles([]); setTargetFormats({}); setError(''); setConvertError('') }

  const handleFormatChange = (file, fmt) => {
    const key = file.name + file.size
    setTargetFormats(prev => ({ ...prev, [key]: fmt }))
  }

  const hasConversions = Object.keys(targetFormats).length > 0

  const handleConvert = async () => {
    setConverting(true)
    setConvertError('')
    const apiUrl = `${supabaseUrl}/functions/v1/convert-file`
    let successCount = 0
    let failCount = 0

    for (const [key, targetFmt] of Object.entries(targetFormats)) {
      const file = files.find(f => f.name + f.size === key)
      if (!file) continue

      try {
        const formData = new FormData()
        formData.append('file', file)
        formData.append('targetFormat', targetFmt)

        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${supabaseAnonKey}`,
          },
          body: formData,
        })

        if (!response.ok) {
          const errData = await response.json().catch(() => ({ error: 'Error desconocido' }))
          throw new Error(errData.error || `Error ${response.status}`)
        }

        const blob = await response.blob()
        const baseName = file.name.replace(/\.[^.]+$/, '')
        const newFileName = `${baseName}${targetFmt}`

        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = newFileName
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        URL.revokeObjectURL(url)

        successCount++
      } catch (err) {
        console.error(`Error converting ${file.name}:`, err)
        failCount++
      }
    }

    setConverting(false)

    if (failCount > 0) {
      setConvertError(`${failCount} archivo${failCount !== 1 ? 's' : ''} no se pudieron convertir. ${successCount} convertido${successCount !== 1 ? 's' : ''} correctamente.`)
    } else {
      setTargetFormats({})
    }
  }

  return (
    <div className="app">
      <header className="app-header">
        <div className="header-logo">
          <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
            <rect width="32" height="32" rx="9" fill="#0f62fe"/>
            <path d="M9 23h14M16 9v10M12 13l4-4 4 4" stroke="#fff" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
        </div>
        <div className="header-text">
          <h1 className="app-title">File Converter</h1>
          <p className="app-subtitle">Detecta y convierte el formato de tus archivos</p>
        </div>
      </header>

      <main className="main">
        <div
          className={`drop-zone${dragging ? ' drop-zone--active' : ''}`}
          onDrop={onDrop}
          onDragOver={onDragOver}
          onDragLeave={onDragLeave}
          onClick={() => inputRef.current.click()}
          role="button"
          tabIndex={0}
          onKeyDown={(e) => e.key === 'Enter' && inputRef.current.click()}
        >
          <input
            ref={inputRef}
            type="file"
            accept={ALL_ACCEPTED}
            multiple
            onChange={onInputChange}
            style={{ display: 'none' }}
          />
          <div className="drop-upload-icon">
            <svg width="40" height="40" viewBox="0 0 40 40" fill="none">
              <path d="M20 28V14M14 20l6-6 6 6" stroke="#0f62fe" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/>
              <path d="M10 30h20" stroke="#0f62fe" strokeWidth="2.5" strokeLinecap="round"/>
            </svg>
          </div>
          <p className="drop-title">
            {dragging ? 'Suelta aquí para subir' : 'Arrastra tus archivos aquí'}
          </p>
          <p className="drop-or">o <span className="drop-link">selecciona desde tu equipo</span></p>
          <div className="format-pills">
            <span className="pill pill--word">DOC / DOCX / ODT</span>
            <span className="pill pill--excel">XLS / XLSX / ODS / CSV</span>
            <span className="pill pill--ppt">PPT / PPTX / ODP</span>
            <span className="pill pill--pdf">PDF</span>
            <span className="pill pill--img">JPG / PNG / SVG</span>
          </div>
        </div>

        {error && (
          <div className="error-banner" role="alert">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="6.5" stroke="#dc2626" strokeWidth="1.5"/>
              <path d="M8 5v3M8 10v1" stroke="#dc2626" strokeWidth="1.5" strokeLinecap="round"/>
            </svg>
            {error}
          </div>
        )}

        {convertError && (
          <div className="error-banner" role="alert">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="6.5" stroke="#dc2626" strokeWidth="1.5"/>
              <path d="M8 5v3M8 10v1" stroke="#dc2626" strokeWidth="1.5" strokeLinecap="round"/>
            </svg>
            {convertError}
          </div>
        )}

        {files.length > 0 && (
          <section className="results">
            <div className="results-header">
              <h2 className="results-title">
                {files.length} archivo{files.length !== 1 ? 's' : ''} detectado{files.length !== 1 ? 's' : ''}
              </h2>
              <button className="btn-clear" onClick={clearAll} disabled={converting}>Limpiar todo</button>
            </div>
            <div className="file-list">
              {files.map((file, i) => (
                <FileCard
                  key={file.name + file.size + i}
                  file={file}
                  onRemove={removeFile}
                  targetFormat={targetFormats[file.name + file.size] || null}
                  onFormatChange={(fmt) => handleFormatChange(file, fmt)}
                  converting={converting}
                />
              ))}
            </div>

            {hasConversions && (
              <div className="convert-summary">
                <div className="convert-summary-title">Conversiones seleccionadas</div>
                <div className="convert-summary-list">
                  {Object.entries(targetFormats).map(([key, fmt]) => {
                    const file = files.find(f => f.name + f.size === key)
                    if (!file) return null
                    const catKey = getCategoryForFile(file)
                    const cat = catKey ? CATEGORIES[catKey] : null
                    return (
                      <div key={key} className="convert-summary-item">
                        <span className="convert-summary-name">{file.name}</span>
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                          <path d="M2 7h10M8 3l4 4-4 4" stroke="#0f62fe" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                        <span className="convert-summary-badge" style={{ background: cat ? cat.color : '#6b7280' }}>
                          {fmt.replace('.', '').toUpperCase()}
                        </span>
                      </div>
                    )
                  })}
                </div>
                <button className="btn-convert" onClick={handleConvert} disabled={converting}>
                  {converting ? (
                    <>
                      <span className="spinner" />
                      Convirtiendo...
                    </>
                  ) : (
                    <>
                      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                        <path d="M2 8h12M10 4l4 4-4 4" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      Convertir y descargar
                    </>
                  )}
                </button>
              </div>
            )}
          </section>
        )}
      </main>
    </div>
  )
}
