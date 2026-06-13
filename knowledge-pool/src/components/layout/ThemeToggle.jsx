import { Sun, Moon } from 'lucide-react'
import { useTheme } from '../../theme/ThemeContext'

function ThemeToggle() {
  const { themeName, toggleTheme } = useTheme()
  const isDark = themeName === 'dark'
  return (
    <button
      type="button"
      className="theme-toggle"
      onClick={toggleTheme}
      role="switch"
      aria-checked={!isDark}
      aria-label={`Switch to ${isDark ? 'light' : 'dark'} theme`}
      title={`Switch to ${isDark ? 'light' : 'dark'} theme`}
    >
      <span className="theme-toggle__thumb">
        <span className="theme-toggle__icon">
          {isDark ? <Moon size={14} /> : <Sun size={14} />}
        </span>
      </span>
    </button>
  )
}

export default ThemeToggle
