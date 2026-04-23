import { startTransition, useEffect, useState } from 'react'
import { DirectoryTable } from './components/DirectoryTable'
import type { AppSettings, AuthState, DirectoryPayload } from './shared/types'

function App() {
  const [authState, setAuthState] = useState<AuthState | null>(null)
  const [payload, setPayload] = useState<DirectoryPayload | null>(null)
  const [loading, setLoading] = useState(true)
  const [savingSettings, setSavingSettings] = useState(false)
  const [showSettings, setShowSettings] = useState(false)
  const [settingsForm, setSettingsForm] = useState<AppSettings>({
    clientId: '',
    tenantId: '',
  })
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let active = true

    async function bootstrap() {
      if (!window.teamspy) {
        if (active) {
          setLoading(false)
          setError('TeamSpy must be launched inside the desktop shell.')
        }
        return
      }

      try {
        const state = await window.teamspy.auth.getState()

        if (!active) {
          return
        }

        setAuthState(state)
        setSettingsForm(state.settings)

        if (state.signedIn) {
          const nextPayload = await window.teamspy.directory.load()
          if (active) {
            setPayload(nextPayload)
          }
        }
      } catch (nextError) {
        if (active) {
          setError(nextError instanceof Error ? nextError.message : 'Unknown error')
        }
      } finally {
        if (active) {
          setLoading(false)
        }
      }
    }

    void bootstrap()

    return () => {
      active = false
    }
  }, [])

  const handleSignIn = () => {
    startTransition(async () => {
      try {
        setLoading(true)
        setError(null)
        const nextState = await window.teamspy.auth.login()
        setAuthState(nextState)

        if (nextState.signedIn) {
          const nextPayload = await window.teamspy.directory.load()
          setPayload(nextPayload)
        }
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : 'Sign-in failed.')
      } finally {
        setLoading(false)
      }
    })
  }

  const handleSignOut = () => {
    startTransition(async () => {
      setLoading(true)
      setError(null)
      const nextState = await window.teamspy.auth.logout()
      setAuthState(nextState)
      setPayload(null)
      setLoading(false)
    })
  }

  const handleSaveSettings = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()

    startTransition(async () => {
      try {
        setSavingSettings(true)
        setError(null)

        const nextState = await window.teamspy.auth.saveSettings(settingsForm)
        setAuthState(nextState)
        setPayload(null)
        setShowSettings(false)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : 'Could not save settings.')
      } finally {
        setSavingSettings(false)
      }
    })
  }

  const handleRefresh = () => {
    startTransition(async () => {
      try {
        setLoading(true)
        setError(null)
        const nextPayload = await window.teamspy.directory.load()
        setPayload(nextPayload)
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : 'Refresh failed.')
      } finally {
        setLoading(false)
      }
    })
  }

  const account = authState?.account

  return (
    <main className="shell">
      <div className="backdrop" />
      <header className="hero-card">
        <div className="hero-copy">
          <p className="eyebrow">Standalone desktop app</p>
          <h1>TeamSpy</h1>
          <p className="lede">
            A local Microsoft Teams directory for macOS and Windows with live presence, sortable
            columns, draggable table layout, and one-click Teams actions.
          </p>
        </div>
        <div className="hero-actions">
          {account ? (
            <>
              <div className="identity-card">
                <span>Connected as</span>
                <strong>{account.name}</strong>
                <small>{account.username}</small>
              </div>
              <button
                className="secondary-button"
                onClick={() => setShowSettings((value) => !value)}
                type="button"
              >
                Settings
              </button>
              <button className="secondary-button" onClick={handleSignOut} type="button">
                Sign out
              </button>
            </>
          ) : (
            <>
              <button
                className="secondary-button"
                onClick={() => setShowSettings((value) => !value)}
                type="button"
              >
                Settings
              </button>
              <button
                className="primary-button"
                disabled={loading || authState?.configured === false}
                onClick={handleSignIn}
                type="button"
              >
                {loading ? 'Connecting…' : 'Connect Microsoft 365'}
              </button>
            </>
          )}
        </div>
      </header>

      {error ? <div className="error-banner">{error}</div> : null}

      {authState?.configured === false || showSettings ? (
        <section className="setup-card">
          <p className="eyebrow">App settings</p>
          <h2>Configure your Microsoft app registration</h2>
          <p>
            Enter the Microsoft Entra application ID and tenant ID directly in TeamSpy. These
            values are stored in the app profile on this machine, so no manual env file is
            required.
          </p>
          <form className="settings-form" onSubmit={handleSaveSettings}>
            <div className="settings-grid">
              <label className="field">
                <span>Client ID</span>
                <input
                  autoComplete="off"
                  onChange={(event) =>
                    setSettingsForm((current) => ({
                      ...current,
                      clientId: event.target.value,
                    }))
                  }
                  placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                  type="text"
                  value={settingsForm.clientId}
                />
              </label>
              <label className="field">
                <span>Tenant ID</span>
                <input
                  autoComplete="off"
                  onChange={(event) =>
                    setSettingsForm((current) => ({
                      ...current,
                      tenantId: event.target.value,
                    }))
                  }
                  placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                  type="text"
                  value={settingsForm.tenantId}
                />
              </label>
            </div>
            <div className="settings-tips">
              <p>Authentication platform: Mobile and desktop applications</p>
              <p>Redirect URI: <code>http://localhost</code></p>
            </div>
            <div className="form-actions">
              <button className="primary-button" disabled={savingSettings} type="submit">
                {savingSettings ? 'Saving…' : 'Save settings'}
              </button>
              {showSettings && authState?.configured ? (
                <button
                  className="secondary-button"
                  onClick={() => setShowSettings(false)}
                  type="button"
                >
                  Close
                </button>
              ) : null}
            </div>
          </form>
          {authState?.missingSettings.length ? (
            <p className="subtle">
              Missing: {authState.missingSettings.join(', ')}
            </p>
          ) : (
            <p className="subtle">
              Changing these values clears the cached Microsoft session for this app.
            </p>
          )}
        </section>
      ) : null}

      {authState?.configured && !authState.signedIn && !loading ? (
        <section className="empty-card">
          <h2>Sign in to load your Teams directory</h2>
          <p>
            TeamSpy runs locally on your Mac. The sign-in flow opens Microsoft in your browser and
            returns to the app with cached tokens.
          </p>
        </section>
      ) : null}

      {authState?.signedIn && payload ? (
        <DirectoryTable
          accountId={account?.homeAccountId ?? 'anonymous'}
          loading={loading}
          onRefresh={handleRefresh}
          payload={payload}
        />
      ) : null}
    </main>
  )
}

export default App
