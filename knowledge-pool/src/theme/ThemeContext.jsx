import { createContext, useContext, useEffect, useLayoutEffect, useState, useCallback } from 'react'
import { themes, themeToCssVars } from './tokens'

const STORAGE_KEY = 'st-portal-theme'
const ThemeContext = createContext({ themeName: 'dark', toggleTheme: () => {}, setThemeName: () => {} })

function getInitialTheme() {
  if (typeof window === 'undefined') return 'dark'
  const saved = window.localStorage.getItem(STORAGE_KEY)
  if (saved === 'dark' || saved === 'light') return saved
  // Respect OS preference on first visit; default dark otherwise.
  if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) return 'light'
  return 'dark'
}

export function ThemeProvider({ children }) {
  const [themeName, setThemeName] = useState(getInitialTheme)

  // Apply tokens as CSS custom properties on the document root.
  useLayoutEffect(() => {
    const theme = themes[themeName] || themes.dark
    const vars = themeToCssVars(theme)
    const root = document.documentElement
    for (const [name, value] of Object.entries(vars)) {
      root.style.setProperty(name, value)
    }
    root.setAttribute('data-theme', themeName)
    root.style.colorScheme = themeName
  }, [themeName])

  useEffect(() => {
    try { window.localStorage.setItem(STORAGE_KEY, themeName) } catch { /* ignore */ }
  }, [themeName])

  const toggleTheme = useCallback(() => {
    setThemeName((t) => (t === 'dark' ? 'light' : 'dark'))
  }, [])

  return (
    <ThemeContext.Provider value={{ themeName, toggleTheme, setThemeName }}>
      {children}
    </ThemeContext.Provider>
  )
}

export function useTheme() {
  return useContext(ThemeContext)
}
