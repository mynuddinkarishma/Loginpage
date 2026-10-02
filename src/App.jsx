import { useEffect, useState } from 'react'
import { ArrowLeft, ArrowRight, Eye, EyeOff, FileSpreadsheet, FileText, Files, Image, Lock, LogOut, Mail, ShieldCheck, Upload, User, Zap } from 'lucide-react'
import { FaGithub } from 'react-icons/fa6'
import './App.css'
import './Workspace.css'
import './Conversion.css'

const API_BASE_URL = (import.meta.env.VITE_API_URL || (import.meta.env.DEV ? 'http://localhost:5001' : window.location.origin)).replace(/\/$/, '')
const conversionTools = [
  { id: 'pdf-to-word', title: 'PDF to Word', format: 'PDF to DOCX', description: 'Extract selectable PDF text into an editable Word file.', accepts: '.pdf,application/pdf', supportedFiles: 'PDF files with selectable text', icon: FileText, tone: 'red' },
  { id: 'word-to-excel', title: 'Word to Excel', format: 'DOCX to XLSX', description: 'Move paragraphs and tables into spreadsheet rows.', accepts: '.docx', supportedFiles: 'Word documents (.docx)', icon: FileSpreadsheet, tone: 'green' },
  { id: 'excel-to-word', title: 'Excel to Word', format: 'XLSX to DOCX', description: 'Export workbook sheets as formatted Word tables.', accepts: '.xlsx', supportedFiles: 'Excel workbooks (.xlsx)', icon: FileSpreadsheet, tone: 'blue' },
  { id: 'image-to-word', title: 'Image to Word', format: 'PNG/JPEG to DOCX', description: 'Recognize text in images and create an editable Word file.', accepts: '.png,.jpg,.jpeg,image/png,image/jpeg', supportedFiles: 'PNG or JPEG images; English OCR', icon: Image, tone: 'orange' },
  { id: 'pdf-tools', title: 'PDF Tools', format: 'Merge or extract pages', description: 'Combine PDFs or save selected pages into a new PDF.', accepts: '.pdf,application/pdf', supportedFiles: 'PDF files', icon: Files, tone: 'charcoal' },
]

export default function App() {
  const [isSignUp, setIsSignUp] = useState(false)
  const [showPassword, setShowPassword] = useState(false)
  const [statusMessage, setStatusMessage] = useState('')
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isAuthenticated, setIsAuthenticated] = useState(false)
  const [authenticatedEmail, setAuthenticatedEmail] = useState('')
  const [selectedToolId, setSelectedToolId] = useState('')
  const [selectedFiles, setSelectedFiles] = useState([])
  const [pdfAction, setPdfAction] = useState('merge')
  const [pageRange, setPageRange] = useState('1')
  const [isConverting, setIsConverting] = useState(false)
  const [conversionMessage, setConversionMessage] = useState('')
  const [conversionError, setConversionError] = useState('')
  const activeTool = conversionTools.find((tool) => tool.id === selectedToolId)

  useEffect(() => {
    fetch(`${API_BASE_URL}/api/traffic/record`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'pageview', path: window.location.pathname }),
    }).catch((error) => console.error('Could not record page view:', error))
  }, [])

  const handleSubmit = async (event) => {
    event.preventDefault()
    const formData = new FormData(event.currentTarget)
    const name = String(formData.get('name') || '').trim()
    const email = String(formData.get('email') || '').trim()
    const password = String(formData.get('password') || '')

    if (isSignUp && password !== formData.get('confirmPassword')) {
      setStatusMessage('Passwords do not match.')
      return
    }

    setIsSubmitting(true)
    setStatusMessage('')

    try {
      const response = await fetch(`${API_BASE_URL}/api/${isSignUp ? 'signup' : 'signin'}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, email, password }),
      })
      const result = await response.json()

      if (!response.ok) {
        throw new Error(result.message || 'Request failed.')
      }

      if (isSignUp) {
        setStatusMessage(result.message)
      } else {
        setAuthenticatedEmail(email)
        setIsAuthenticated(true)
      }
    } catch (error) {
      setStatusMessage(error.message || 'Could not connect to the server.')
    } finally {
      setIsSubmitting(false)
    }
  }

  const handleConversionSubmit = async (event) => {
    event.preventDefault()
    if (!activeTool) return

    const isPdfTools = activeTool.id === 'pdf-tools'
    const endpoint = isPdfTools ? `/api/pdf-tools/${pdfAction}` : `/api/convert/${activeTool.id}`
    const fieldName = isPdfTools && pdfAction === 'merge' ? 'files' : 'file'
    const formData = new FormData()
    selectedFiles.forEach((file) => formData.append(fieldName, file))
    if (isPdfTools && pdfAction === 'extract') formData.append('pages', pageRange)

    setIsConverting(true)
    setConversionError('')
    setConversionMessage('')

    try {
      const response = await fetch(`${API_BASE_URL}${endpoint}`, { method: 'POST', body: formData })
      if (!response.ok) {
        const result = await response.json().catch(() => null)
        throw new Error(result?.message || 'The file could not be converted.')
      }

      const blob = await response.blob()
      const disposition = response.headers.get('content-disposition') || ''
      const filename = disposition.match(/filename="?([^";]+)"?/i)?.[1] || 'converted-file'
      const downloadUrl = URL.createObjectURL(blob)
      const downloadLink = document.createElement('a')
      downloadLink.href = downloadUrl
      downloadLink.download = filename
      document.body.append(downloadLink)
      downloadLink.click()
      downloadLink.remove()
      setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000)
      setConversionMessage(`${filename} is ready.`)
    } catch (error) {
      setConversionError(error.message || 'Could not connect to the conversion server.')
    } finally {
      setIsConverting(false)
    }
  }

  const openTool = (toolId) => {
    setSelectedToolId(toolId)
    setSelectedFiles([])
    setPdfAction('merge')
    setPageRange('1')
    setConversionMessage('')
    setConversionError('')
  }

  const returnToTools = () => {
    setSelectedToolId('')
    setSelectedFiles([])
    setConversionMessage('')
    setConversionError('')
  }

  if (isAuthenticated) {
    return (
      <div className="workspace-page">
        <div className="workspace-shell">
          <header className="workspace-header">
            <div className="workspace-brand"><span className="workspace-brand-mark"><FileText size={24} /></span><span>PDFFlow</span></div>
            <div className="workspace-account"><span className="workspace-email">{authenticatedEmail}</span><button type="button" className="sign-out-button" onClick={() => { setIsAuthenticated(false); returnToTools() }}><LogOut size={17} /> Sign out</button></div>
          </header>
          {activeTool ? (
            <main className="workspace-content converter-content">
              <button type="button" className="workspace-back-button" onClick={returnToTools}><ArrowLeft size={17} /> All tools</button>
              <div className="workspace-heading"><p className="workspace-eyebrow">{activeTool.format}</p><h1>{activeTool.title}</h1><p>{activeTool.description}</p></div>
              <form className="converter-panel" onSubmit={handleConversionSubmit}>
                {activeTool.id === 'pdf-tools' && (
                  <div className="pdf-action-switch" aria-label="PDF action">
                    <button type="button" className={pdfAction === 'merge' ? 'active' : ''} onClick={() => { setPdfAction('merge'); setSelectedFiles([]); setConversionError(''); setConversionMessage('') }}>Merge PDFs</button>
                    <button type="button" className={pdfAction === 'extract' ? 'active' : ''} onClick={() => { setPdfAction('extract'); setSelectedFiles([]); setConversionError(''); setConversionMessage('') }}>Extract pages</button>
                  </div>
                )}
                <label className="upload-zone" htmlFor="conversion-files">
                  <input
                    id="conversion-files"
                    className="upload-input"
                    type="file"
                    accept={activeTool.accepts}
                    multiple={activeTool.id === 'pdf-tools' && pdfAction === 'merge'}
                    onChange={(event) => { setSelectedFiles(Array.from(event.target.files || [])); setConversionError(''); setConversionMessage(''); event.target.value = '' }}
                  />
                  <span className="upload-icon"><Upload size={23} /></span>
                  <strong>{selectedFiles.length ? `${selectedFiles.length} file${selectedFiles.length === 1 ? '' : 's'} selected` : 'Choose file'}</strong>
                  <span>{activeTool.id === 'pdf-tools' && pdfAction === 'merge' ? 'Select 2 to 5 PDF files' : activeTool.supportedFiles}</span>
                </label>
                {selectedFiles.length > 0 && (
                  <ul className="selected-files" aria-label="Selected files">
                    {selectedFiles.map((file, index) => <li key={`${file.name}-${index}`}><span>{file.name}</span><span>{(file.size / 1024 / 1024).toFixed(2)} MB</span></li>)}
                  </ul>
                )}
                {activeTool.id === 'pdf-tools' && pdfAction === 'extract' && (
                  <label className="page-range-field" htmlFor="page-range">Pages to extract<input id="page-range" value={pageRange} onChange={(event) => setPageRange(event.target.value)} placeholder="1-3, 5" required /></label>
                )}
                <button
                  type="submit"
                  className="convert-button"
                  disabled={isConverting || (activeTool.id === 'pdf-tools' && pdfAction === 'merge' ? selectedFiles.length < 2 : selectedFiles.length !== 1) || (activeTool.id === 'pdf-tools' && pdfAction === 'extract' && !pageRange.trim())}
                >
                  {isConverting ? 'Working...' : activeTool.id === 'pdf-tools' ? pdfAction === 'merge' ? 'Merge and download' : 'Extract and download' : 'Convert and download'}
                  {isConverting ? null : <ArrowRight size={18} />}
                </button>
                {conversionError && <p className="conversion-feedback error" role="alert">{conversionError}</p>}
                {conversionMessage && <p className="conversion-feedback success" role="status">{conversionMessage}</p>}
              </form>
            </main>
          ) : (
            <main className="workspace-content">
              <div className="workspace-heading"><p className="workspace-eyebrow">YOUR WORKSPACE</p><h1>Choose a document tool</h1><p>Pick a workflow to get started with your files.</p></div>
              <section className="workspace-tools" aria-label="Document tools">
                {conversionTools.map(({ id, title, format, description, icon: ToolIcon, tone }) => (
                  <button type="button" className={`workspace-tool ${tone}`} key={id} onClick={() => openTool(id)}>
                    <span className="workspace-tool-icon"><ToolIcon size={23} /></span>
                    <span className="workspace-tool-format">{format}</span>
                    <span className="workspace-tool-title">{title}</span>
                    <span className="workspace-tool-description">{description}</span>
                    <ArrowRight className="workspace-tool-arrow" size={18} />
                  </button>
                ))}
              </section>
            </main>
          )}
        </div>
      </div>
    )
  }

  return (
    <div className="auth-page">
      <div className="auth-shell">
        <aside className="brand-panel">
          <div>
            <div className="brand-header"><div className="brand-mark"><FileText size={28} /></div><span className="brand-name">PDFFlow</span></div>
            <div className="brand-copy"><h1>Master Your Documents.</h1></div>
            <p className="brand-description">Convert, merge, split, and edit PDFs with lightning speed and military-grade security. Join millions of professionals today.</p>
            <div className="feature-list">
              <div className="feature-item"><div className="feature-icon"><FileText size={24} /></div><div><h4>Lossless Conversion</h4><p>Keep your formatting intact.</p></div></div>
              <div className="feature-item"><div className="feature-icon"><Zap size={24} /></div><div><h4>Lightning Fast</h4><p>Process large files in seconds.</p></div></div>
              <div className="feature-item"><div className="feature-icon"><ShieldCheck size={24} /></div><div><h4>Secure &amp; Private</h4><p>256-bit encryption for all files.</p></div></div>
            </div>
          </div>
          <div className="brand-footer">© 2026 PDFFlow Inc. All rights reserved.</div>
        </aside>
        <main className="form-panel">
          <div className="form-card">
            <div className="mode-toggle"><div className={`toggle-slider ${isSignUp ? 'right' : ''}`} /><button type="button" className={!isSignUp ? 'toggle-button active' : 'toggle-button'} onClick={() => setIsSignUp(false)}>Sign In</button><button type="button" className={isSignUp ? 'toggle-button active' : 'toggle-button'} onClick={() => setIsSignUp(true)}>Sign Up</button></div>
            <div className="form-header"><h2>{isSignUp ? 'Create an account' : 'Welcome back'}</h2><p>{isSignUp ? 'Start converting PDFs for free today.' : 'Enter your details to access your dashboard.'}</p></div>
            <form className="auth-form" onSubmit={handleSubmit}>
              {isSignUp && <div className="field-group"><label className="field-label" htmlFor="name">Full Name</label><div className="input-wrap"><User size={20} className="field-icon" /><input id="name" name="name" placeholder="Full Name" required /></div></div>}
              <div className="field-group"><label className="field-label" htmlFor="email">Email Address</label><div className="input-wrap"><Mail size={20} className="field-icon" /><input id="email" name="email" type="email" placeholder="Email Address" required /></div></div>
              <div className="field-group"><div className="password-header"><label className="field-label" htmlFor="password">Password</label>{!isSignUp && <a href="#" className="link-text">Forgot password?</a>}</div><div className="input-wrap"><Lock size={20} className="field-icon" /><input id="password" name="password" type={showPassword ? 'text' : 'password'} placeholder="Password" required /><button type="button" className="password-toggle" onClick={() => setShowPassword((value) => !value)} aria-label="Toggle password visibility">{showPassword ? <EyeOff size={18} /> : <Eye size={18} />}</button></div></div>
              {isSignUp && <div className="field-group"><label className="field-label" htmlFor="confirm-password">Confirm Password</label><div className="input-wrap"><Lock size={20} className="field-icon" /><input id="confirm-password" name="confirmPassword" type={showPassword ? 'text' : 'password'} placeholder="Confirm Password" required /></div></div>}
              {!isSignUp && <div className="remember-row"><label className="remember-me"><input type="checkbox" /><span>Remember me</span></label><a href="#" className="link-text">Forgot password?</a></div>}
              <button type="submit" className="primary-button" disabled={isSubmitting}>{isSubmitting ? 'Please wait...' : isSignUp ? 'Create Account' : 'Sign In'} <ArrowRight size={18} /></button>
              {statusMessage && <p role="status" aria-live="polite">{statusMessage}</p>}
            </form>
            <div className="divider"><span>Or continue with</span></div>
            <div className="social-row"><button type="button" className="social-button"><span className="google-logo">G</span>Google</button><button type="button" className="social-button"><FaGithub size={18} />GitHub</button></div>
            <div className="switch-footer"><span>{isSignUp ? 'Already have an account?' : "Don't have an account?"}</span> <button type="button" onClick={() => setIsSignUp((value) => !value)}>{isSignUp ? 'Sign In' : 'Sign up for free'}</button></div>
          </div>
        </main>
      </div>
    </div>
  )
}