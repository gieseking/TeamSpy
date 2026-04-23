import { startTransition, useEffect, useState } from 'react'
import { DirectoryTable } from './components/DirectoryTable'
import type { AuthState, DirectoryPayload } from './shared/types'

function App() {
  const [authState, setAuthState] = useState<AuthState | null>(null)
  const [payload, setPayload] = useState<DirectoryPayload | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  const loadDirectory = async () => {
    const nextPayload = await window.friendlyfaces.directory.load()
    setPayload(nextPayload)
    return nextPayload
  }

  const connectAndLoad = async () => {
    const nextState = await window.friendlyfaces.auth.login()
    setAuthState(nextState)

    if (nextState.signedIn) {
      await loadDirectory()
    }

    return nextState
  }

  useEffect(() => {
    let active = true

    async function bootstrap() {
      if (!window.friendlyfaces) {
        if (active) {
          setLoading(false)
          setError('FriendlyFaces must be launched inside the desktop shell.')
        }
        return
      }

      try {
        let state = await window.friendlyfaces.auth.getState()

        if (!active) {
          return
        }

        setAuthState(state)

        if (state.configured && !state.signedIn) {
          state = await window.friendlyfaces.auth.login()

          if (!active) {
            return
          }

          setAuthState(state)
        }

        if (state.signedIn) {
          const nextPayload = await window.friendlyfaces.directory.load()

          if (!active) {
            return
          }

          setPayload(nextPayload)
        }
      } catch (nextError) {
        if (active) {
          setPayload(null)
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
        await connectAndLoad()
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
      const nextState = await window.friendlyfaces.auth.logout()
      setAuthState(nextState)
      setPayload(null)
      setLoading(false)
    })
  }

  const handleRefresh = () => {
    startTransition(async () => {
      try {
        setLoading(true)
        setError(null)
        await loadDirectory()
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
          <h1>FriendlyFaces</h1>
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
              <button className="secondary-button" onClick={handleSignOut} type="button">
                Sign out
              </button>
            </>
          ) : (
            <button
              className="primary-button"
              disabled={loading || authState?.configured === false}
              onClick={handleSignIn}
              type="button"
            >
              {loading ? 'Connecting…' : 'Connect Microsoft 365'}
            </button>
          )}
        </div>
      </header>

      {error ? <div className="error-banner">{error}</div> : null}

      {authState?.configured === false ? (
        <section className="setup-card">
          <p className="eyebrow">Publisher setup</p>
          <h2>This build is missing a Microsoft client ID</h2>
          <p>
            End users should not have to enter tenant IDs or app registration values. FriendlyFaces now
            resolves the tenant from the Microsoft 365 sign-in and only needs a publisher-supplied
            public client ID.
          </p>
          <div className="settings-tips">
            <p>Authority: <code>https://login.microsoftonline.com/organizations</code></p>
            <p>Redirect URI: <code>http://localhost</code></p>
            <p>Missing: {authState.missingConfiguration.join(', ')}</p>
          </div>
          <p className="subtle">
            To finish this build, set the public client ID in
            <code> electron/publisher-config.ts</code>.
          </p>
        </section>
      ) : null}

      {authState?.configured && !authState.signedIn && !loading ? (
        <section className="empty-card">
          <h2>Sign in to continue</h2>
          <p>
            FriendlyFaces starts by opening Microsoft 365 authentication inside the app. If you closed
            that dialog, use the button above to reconnect and continue to the grid.
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
