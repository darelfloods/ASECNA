import React, { useState } from 'react';
import { login, register, LoginCredentials, RegisterData, AuthResponse, User } from '../services/authService';

interface AuthProps {
  onAuthSuccess: (user: User) => void;
}

type AuthMode = 'login' | 'register';

export const Auth: React.FC<AuthProps> = ({ onAuthSuccess }) => {
  const [mode, setMode] = useState<AuthMode>('login');
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [registrationPending, setRegistrationPending] = useState(false);
  const [registeredEmail, setRegisteredEmail] = useState<string>('');

  // Champs du formulaire
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [nom, setNom] = useState('');
  const [prenom, setPrenom] = useState('');
  const [matricule, setMatricule] = useState('');
  const [service, setService] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  const resetForm = () => {
    setEmail('');
    setPassword('');
    setConfirmPassword('');
    setNom('');
    setPrenom('');
    setMatricule('');
    setService('');
    setError(null);
    setSuccess(null);
  };

  const switchMode = (newMode: AuthMode) => {
    resetForm();
    setRegistrationPending(false);
    setRegisteredEmail('');
    setMode(newMode);
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoading(true);
    setError(null);

    try {
      const credentials: LoginCredentials = { email, password };
      const response: AuthResponse = login(credentials);

      if (response.success && response.user) {
        setSuccess('Connexion réussie ! Redirection...');
        setTimeout(() => {
          onAuthSuccess(response.user!);
        }, 1000);
      } else {
        setError(response.message);
      }
    } catch (err) {
      setError('Une erreur est survenue. Veuillez réessayer.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoading(true);
    setError(null);

    try {
      const data: RegisterData = {
        email,
        password,
        confirmPassword,
        nom,
        prenom,
        matricule: matricule || undefined,
        service: service || undefined
      };
      
      const response: AuthResponse = register(data);

      if (response.success) {
        setRegisteredEmail(email);
        setRegistrationPending(true);
        resetForm();
      } else {
        setError(response.message);
      }
    } catch (err) {
      setError('Une erreur est survenue. Veuillez réessayer.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="auth-container">
      <div className="auth-card">
        {/* Logo et titre */}
        <div className="auth-header">
          <div className="auth-logo">
            <img src="/75664_O.jpg" alt="ASECNA Logo" className="auth-logo-img" />
          </div>
          <h1 className="auth-title">ASECNA</h1>
          <p className="auth-subtitle">Service Budget et Facturation</p>
        </div>

        {/* Onglets Login/Register */}
        <div className="auth-tabs">
          <button
            className={`auth-tab ${mode === 'login' ? 'active' : ''}`}
            onClick={() => switchMode('login')}
          >
            Connexion
          </button>
          <button
            className={`auth-tab ${mode === 'register' ? 'active' : ''}`}
            onClick={() => switchMode('register')}
          >
            Inscription
          </button>
        </div>

        {/* Messages d'erreur/succès */}
        {error && (
          <div className="auth-message auth-error">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" />
              <line x1="15" y1="9" x2="9" y2="15" />
              <line x1="9" y1="9" x2="15" y2="15" />
            </svg>
            {error}
          </div>
        )}
        {success && (
          <div className="auth-message auth-success">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
              <polyline points="22,4 12,14.01 9,11.01" />
            </svg>
            {success}
          </div>
        )}

        {/* Formulaire de connexion */}
        {mode === 'login' && (
          <form onSubmit={handleLogin} className="auth-form">
            <div className="auth-field">
              <label htmlFor="login-email">Email</label>
              <input
                id="login-email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Entrer votre email"
                required
                autoComplete="email"
              />
            </div>

            <div className="auth-field">
              <label htmlFor="login-password">Mot de passe</label>
              <div className="auth-password-wrapper">
                <input
                  id="login-password"
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  required
                  autoComplete="current-password"
                />
                <button
                  type="button"
                  className="auth-password-toggle"
                  onClick={() => setShowPassword(!showPassword)}
                >
                  {showPassword ? (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                      <line x1="1" y1="1" x2="23" y2="23" />
                    </svg>
                  ) : (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                      <circle cx="12" cy="12" r="3" />
                    </svg>
                  )}
                </button>
              </div>
            </div>

            <button type="submit" className="auth-submit" disabled={isLoading}>
              {isLoading ? (
                <>
                  <span className="auth-spinner"></span>
                  Connexion...
                </>
              ) : (
                'Se connecter'
              )}
            </button>

          </form>
        )}

        {/* Écran de confirmation d'inscription en attente */}
        {mode === 'register' && registrationPending && (
          <div className="auth-registration-pending">
            <div className="auth-pending-icon">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" width="56" height="56">
                <circle cx="12" cy="12" r="10" />
                <polyline points="12,6 12,12 16,14" />
              </svg>
            </div>
            <h3 className="auth-pending-title">Inscription enregistrée !</h3>
            <p className="auth-pending-desc">
              Votre demande a bien été reçue. Votre compte est actuellement
              <strong> en attente de validation</strong> par un administrateur.
            </p>
            {registeredEmail && (
              <div className="auth-pending-email">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                  <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" />
                  <polyline points="22,6 12,13 2,6" />
                </svg>
                <span>{registeredEmail}</span>
              </div>
            )}
            <p className="auth-pending-info">
              Vous recevrez un email de confirmation dès que votre compte sera activé.
            </p>
            <button
              type="button"
              className="auth-submit"
              onClick={() => switchMode('login')}
            >
              Retour à la connexion
            </button>
          </div>
        )}

        {/* Formulaire d'inscription */}
        {mode === 'register' && !registrationPending && (
          <form onSubmit={handleRegister} className="auth-form">
            <div className="auth-row">
              <div className="auth-field">
                <label htmlFor="register-nom">Nom *</label>
                <input
                  id="register-nom"
                  type="text"
                  value={nom}
                  onChange={(e) => setNom(e.target.value)}
                  placeholder="Saisissez votre nom"
                  required
                />
              </div>
              <div className="auth-field">
                <label htmlFor="register-prenom">Prénom *</label>
                <input
                  id="register-prenom"
                  type="text"
                  value={prenom}
                  onChange={(e) => setPrenom(e.target.value)}
                  placeholder="Saisissez votre prénom"
                  required
                />
              </div>
            </div>

            <div className="auth-field">
              <label htmlFor="register-email">Email *</label>
              <input
                id="register-email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Entrer votre email professionnel"
                required
                autoComplete="email"
              />
            </div>

            <div className="auth-row">
              <div className="auth-field">
                <label htmlFor="register-matricule">Matricule</label>
                <input
                  id="register-matricule"
                  type="text"
                  value={matricule}
                  onChange={(e) => setMatricule(e.target.value)}
                  placeholder="Matricule (optionnel)"
                />
              </div>
              <div className="auth-field">
                <label htmlFor="register-service">Service</label>
                <input
                  id="register-service"
                  type="text"
                  value={service}
                  onChange={(e) => setService(e.target.value)}
                  placeholder="Service (optionnel)"
                />
              </div>
            </div>

            <div className="auth-field">
              <label htmlFor="register-password">Mot de passe *</label>
              <div className="auth-password-wrapper">
                <input
                  id="register-password"
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Minimum 8 caractères"
                  required
                  minLength={8}
                  autoComplete="new-password"
                />
                <button
                  type="button"
                  className="auth-password-toggle"
                  onClick={() => setShowPassword(!showPassword)}
                >
                  {showPassword ? (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24" />
                      <line x1="1" y1="1" x2="23" y2="23" />
                    </svg>
                  ) : (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                      <circle cx="12" cy="12" r="3" />
                    </svg>
                  )}
                </button>
              </div>
            </div>

            <div className="auth-field">
              <label htmlFor="register-confirm">Confirmer le mot de passe *</label>
              <input
                id="register-confirm"
                type="password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="••••••••"
                required
                autoComplete="new-password"
              />
            </div>

            <button type="submit" className="auth-submit" disabled={isLoading}>
              {isLoading ? (
                <>
                  <span className="auth-spinner"></span>
                  Inscription...
                </>
              ) : (
                "S'inscrire"
              )}
            </button>
          </form>
        )}

        {/* Footer */}
        <div className="auth-footer">
          <p>© {new Date().getFullYear()} ASECNA - Délégation du Gabon</p>
          <p>Usage interne uniquement</p>
        </div>
      </div>
    </div>
  );
};

export default Auth;
