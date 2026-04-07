/**
 * Service d'authentification pour ASECNA Budget et Facturation
 * Toutes les opérations utilisateur passent par l'API serveur (SQLite)
 * La session locale (localStorage) ne stocke que l'utilisateur connecté
 */

export interface User {
  id: string;
  email: string;
  nom: string;
  prenom: string;
  role: 'admin' | 'user' | 'viewer';
  status: 'pending' | 'approved' | 'rejected';
  matricule?: string;
  service?: string;
  createdAt: string;
  lastLogin?: string;
  approvedBy?: string;
  approvedAt?: string;
}

export interface AuthState {
  isAuthenticated: boolean;
  user: User | null;
  token: string | null;
}

export interface LoginCredentials {
  email: string;
  password: string;
}

export interface RegisterData {
  email: string;
  password: string;
  confirmPassword: string;
  nom: string;
  prenom: string;
  matricule?: string;
  service?: string;
}

export interface AuthResponse {
  success: boolean;
  message: string;
  user?: User;
  token?: string;
}

const STORAGE_KEYS = {
  USER: 'asecna_user',
  TOKEN: 'asecna_token'
};

const API_BASE = window.location.hostname === 'localhost'
  ? 'http://localhost:3002/api'
  : '/api';

async function apiFetch(path: string, options?: RequestInit): Promise<any> {
  const res = await fetch(`${API_BASE}${path}`, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });
  return res.json();
}

/**
 * Connexion utilisateur
 */
export async function login(credentials: LoginCredentials): Promise<AuthResponse> {
  const { email, password } = credentials;
  if (!email || !password) {
    return { success: false, message: 'Email et mot de passe requis' };
  }
  const data = await apiFetch('/auth/login', {
    method: 'POST',
    body: JSON.stringify({ email, password })
  });
  if (data.success && data.user && data.token) {
    localStorage.setItem(STORAGE_KEYS.USER, JSON.stringify(data.user));
    localStorage.setItem(STORAGE_KEYS.TOKEN, data.token);
  }
  return data;
}

/**
 * Inscription d'un nouvel utilisateur
 */
export async function register(data: RegisterData): Promise<AuthResponse> {
  const { email, password, confirmPassword, nom, prenom, matricule, service } = data;
  if (!email || !password || !nom || !prenom) {
    return { success: false, message: 'Tous les champs obligatoires doivent être remplis' };
  }
  if (password !== confirmPassword) {
    return { success: false, message: 'Les mots de passe ne correspondent pas' };
  }
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    return { success: false, message: "Format d'email invalide" };
  }
  return apiFetch('/auth/register', {
    method: 'POST',
    body: JSON.stringify({ email, password, nom, prenom, matricule, service })
  });
}

/**
 * Déconnexion
 */
export function logout(): void {
  localStorage.removeItem(STORAGE_KEYS.USER);
  localStorage.removeItem(STORAGE_KEYS.TOKEN);
}

/**
 * Récupère l'état d'authentification actuel (session locale)
 */
export function getAuthState(): AuthState {
  try {
    const userJson = localStorage.getItem(STORAGE_KEYS.USER);
    const token = localStorage.getItem(STORAGE_KEYS.TOKEN);
    if (userJson && token) {
      return { isAuthenticated: true, user: JSON.parse(userJson) as User, token };
    }
  } catch (e) {
    console.error('Erreur lecture état auth:', e);
  }
  return { isAuthenticated: false, user: null, token: null };
}

export function isAuthenticated(): boolean {
  return getAuthState().isAuthenticated;
}

export function getCurrentUser(): User | null {
  return getAuthState().user;
}

/**
 * Met à jour le profil utilisateur
 */
export async function updateProfile(updates: Partial<User>): Promise<AuthResponse> {
  const state = getAuthState();
  if (!state.isAuthenticated || !state.user) {
    return { success: false, message: 'Non authentifié' };
  }
  const data = await apiFetch(`/auth/profile/${state.user.id}`, {
    method: 'PUT',
    body: JSON.stringify(updates)
  });
  if (data.success && data.user) {
    localStorage.setItem(STORAGE_KEYS.USER, JSON.stringify(data.user));
  }
  return data;
}

/**
 * Récupère la liste des utilisateurs en attente (admin)
 */
export async function getPendingUsers(): Promise<User[]> {
  const data = await apiFetch('/auth/users');
  if (!data.success) return [];
  return (data.users as User[]).filter(u => u.status === 'pending')
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
}

/**
 * Récupère tous les utilisateurs (admin)
 */
export async function getAllUsers(): Promise<User[]> {
  const data = await apiFetch('/auth/users');
  if (!data.success) return [];
  return (data.users as User[])
    .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
}

/**
 * Approuve une inscription (admin)
 */
export async function approveUser(userId: string): Promise<AuthResponse> {
  const state = getAuthState();
  return apiFetch(`/auth/approve/${userId}`, {
    method: 'POST',
    body: JSON.stringify({ adminEmail: state.user?.email })
  });
}

/**
 * Change le rôle d'un utilisateur (admin)
 */
export async function changeUserRole(userId: string, newRole: User['role']): Promise<AuthResponse> {
  return apiFetch(`/auth/role/${userId}`, {
    method: 'POST',
    body: JSON.stringify({ role: newRole })
  });
}

/**
 * Rejette une inscription (admin)
 */
export async function rejectUser(userId: string): Promise<AuthResponse> {
  return apiFetch(`/auth/reject/${userId}`, { method: 'POST' });
}

/**
 * Supprime un utilisateur (admin)
 */
export async function deleteUser(userId: string): Promise<AuthResponse> {
  return apiFetch(`/auth/users/${userId}`, { method: 'DELETE' });
}

/**
 * Change le mot de passe
 */
export async function changePassword(currentPassword: string, newPassword: string): Promise<AuthResponse> {
  const state = getAuthState();
  if (!state.isAuthenticated || !state.user) {
    return { success: false, message: 'Non authentifié' };
  }
  return apiFetch('/auth/change-password', {
    method: 'POST',
    body: JSON.stringify({ email: state.user.email, currentPassword, newPassword })
  });
}
