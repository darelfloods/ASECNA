/**
 * Service d'authentification pour ASECNA Budget et Facturation
 * Gère la connexion, l'inscription et la session utilisateur
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

// Clés de stockage local
const STORAGE_KEYS = {
  USER: 'asecna_user',
  TOKEN: 'asecna_token',
  USERS_DB: 'asecna_users_db'
};

// Utilisateur admin par défaut
const DEFAULT_ADMIN: User = {
  id: 'admin-001',
  email: 'admin@asecna.ga',
  nom: 'ADMINISTRATEUR',
  prenom: 'ASECNA',
  role: 'admin',
  status: 'approved',
  matricule: '000001',
  service: 'Direction Générale',
  createdAt: new Date().toISOString()
};

const DEFAULT_ADMIN_PASSWORD = 'Admin@2024';

/**
 * Génère un ID unique pour les utilisateurs
 */
function generateId(): string {
  return 'user-' + Date.now().toString(36) + Math.random().toString(36).substr(2, 9);
}

/**
 * Hash simple du mot de passe (pour demo - en production utiliser bcrypt côté serveur)
 */
function hashPassword(password: string): string {
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const char = password.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return 'hash_' + Math.abs(hash).toString(16) + '_' + password.length;
}

/**
 * Récupère la base de données des utilisateurs depuis le localStorage
 */
function getUsersDB(): Map<string, { user: User; passwordHash: string }> {
  try {
    const data = localStorage.getItem(STORAGE_KEYS.USERS_DB);
    if (data) {
      const parsed = JSON.parse(data);
      return new Map(Object.entries(parsed));
    }
  } catch (e) {
    console.error('Erreur lecture base utilisateurs:', e);
  }
  
  // Initialiser avec l'admin par défaut
  const defaultDB = new Map<string, { user: User; passwordHash: string }>();
  defaultDB.set(DEFAULT_ADMIN.email.toLowerCase(), {
    user: DEFAULT_ADMIN,
    passwordHash: hashPassword(DEFAULT_ADMIN_PASSWORD)
  });
  saveUsersDB(defaultDB);
  return defaultDB;
}

/**
 * Sauvegarde la base de données des utilisateurs
 */
function saveUsersDB(db: Map<string, { user: User; passwordHash: string }>): void {
  try {
    const obj = Object.fromEntries(db);
    localStorage.setItem(STORAGE_KEYS.USERS_DB, JSON.stringify(obj));
  } catch (e) {
    console.error('Erreur sauvegarde base utilisateurs:', e);
  }
}

/**
 * Connexion utilisateur
 */
export function login(credentials: LoginCredentials): AuthResponse {
  const { email, password } = credentials;
  
  if (!email || !password) {
    return { success: false, message: 'Email et mot de passe requis' };
  }
  
  const db = getUsersDB();
  const userRecord = db.get(email.toLowerCase());
  
  if (!userRecord) {
    return { success: false, message: 'Email ou mot de passe incorrect' };
  }
  
  const passwordHash = hashPassword(password);
  if (userRecord.passwordHash !== passwordHash) {
    return { success: false, message: 'Email ou mot de passe incorrect' };
  }
  
  // Vérifier le statut de l'utilisateur
  if (userRecord.user.status === 'pending') {
    return { success: false, message: 'Votre compte est en attente de validation par un administrateur' };
  }
  
  if (userRecord.user.status === 'rejected') {
    return { success: false, message: 'Votre demande d\'inscription a été refusée' };
  }
  
  // Mettre à jour la dernière connexion
  userRecord.user.lastLogin = new Date().toISOString();
  db.set(email.toLowerCase(), userRecord);
  saveUsersDB(db);
  
  // Générer un token simple (en production, utiliser JWT côté serveur)
  const token = 'token_' + generateId() + '_' + Date.now();
  
  // Sauvegarder la session
  localStorage.setItem(STORAGE_KEYS.USER, JSON.stringify(userRecord.user));
  localStorage.setItem(STORAGE_KEYS.TOKEN, token);
  
  return {
    success: true,
    message: 'Connexion réussie',
    user: userRecord.user,
    token
  };
}

/**
 * Inscription d'un nouvel utilisateur
 */
export function register(data: RegisterData): AuthResponse {
  const { email, password, confirmPassword, nom, prenom, matricule, service } = data;
  
  // Validations
  if (!email || !password || !nom || !prenom) {
    return { success: false, message: 'Tous les champs obligatoires doivent être remplis' };
  }
  
  if (password !== confirmPassword) {
    return { success: false, message: 'Les mots de passe ne correspondent pas' };
  }
  
  if (password.length < 8) {
    return { success: false, message: 'Le mot de passe doit contenir au moins 8 caractères' };
  }
  
  // Validation email
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) {
    return { success: false, message: 'Format d\'email invalide' };
  }
  
  const db = getUsersDB();
  
  // Vérifier si l'email existe déjà
  if (db.has(email.toLowerCase())) {
    return { success: false, message: 'Cet email est déjà utilisé' };
  }
  
  // Créer le nouvel utilisateur avec statut "pending"
  const newUser: User = {
    id: generateId(),
    email: email.toLowerCase(),
    nom: nom.toUpperCase(),
    prenom: prenom,
    role: 'user', // Par défaut, rôle utilisateur
    status: 'pending', // En attente de validation par admin
    matricule,
    service,
    createdAt: new Date().toISOString()
  };
  
  // Sauvegarder dans la base
  db.set(email.toLowerCase(), {
    user: newUser,
    passwordHash: hashPassword(password)
  });
  saveUsersDB(db);
  
  // Ne pas connecter automatiquement - attendre validation admin
  return {
    success: true,
    message: 'Inscription enregistrée. Votre compte sera activé après validation par un administrateur.',
    user: newUser
  };
}

/**
 * Déconnexion
 */
export function logout(): void {
  localStorage.removeItem(STORAGE_KEYS.USER);
  localStorage.removeItem(STORAGE_KEYS.TOKEN);
}

/**
 * Récupère l'état d'authentification actuel
 */
export function getAuthState(): AuthState {
  try {
    const userJson = localStorage.getItem(STORAGE_KEYS.USER);
    const token = localStorage.getItem(STORAGE_KEYS.TOKEN);
    
    if (userJson && token) {
      const user = JSON.parse(userJson) as User;
      return {
        isAuthenticated: true,
        user,
        token
      };
    }
  } catch (e) {
    console.error('Erreur lecture état auth:', e);
  }
  
  return {
    isAuthenticated: false,
    user: null,
    token: null
  };
}

/**
 * Vérifie si l'utilisateur est authentifié
 */
export function isAuthenticated(): boolean {
  return getAuthState().isAuthenticated;
}

/**
 * Récupère l'utilisateur courant
 */
export function getCurrentUser(): User | null {
  return getAuthState().user;
}

/**
 * Met à jour le profil utilisateur
 */
export function updateProfile(updates: Partial<User>): AuthResponse {
  const state = getAuthState();
  
  if (!state.isAuthenticated || !state.user) {
    return { success: false, message: 'Non authentifié' };
  }
  
  const db = getUsersDB();
  const userRecord = db.get(state.user.email);
  
  if (!userRecord) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }
  
  // Mettre à jour les champs autorisés
  const updatedUser: User = {
    ...userRecord.user,
    nom: updates.nom || userRecord.user.nom,
    prenom: updates.prenom || userRecord.user.prenom,
    matricule: updates.matricule || userRecord.user.matricule,
    service: updates.service || userRecord.user.service
  };
  
  db.set(state.user.email, { ...userRecord, user: updatedUser });
  saveUsersDB(db);
  
  localStorage.setItem(STORAGE_KEYS.USER, JSON.stringify(updatedUser));
  
  return {
    success: true,
    message: 'Profil mis à jour',
    user: updatedUser
  };
}

/**
 * Récupère la liste des utilisateurs en attente de validation (admin seulement)
 */
export function getPendingUsers(): User[] {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return [];
  }
  
  const db = getUsersDB();
  const pendingUsers: User[] = [];
  
  db.forEach((record) => {
    if (record.user.status === 'pending') {
      pendingUsers.push(record.user);
    }
  });
  
  return pendingUsers.sort((a, b) => 
    new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
  );
}

/**
 * Récupère tous les utilisateurs (admin seulement)
 */
export function getAllUsers(): User[] {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return [];
  }
  
  const db = getUsersDB();
  const users: User[] = [];
  
  db.forEach((record) => {
    users.push(record.user);
  });
  
  return users.sort((a, b) => 
    new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
  );
}

/**
 * Approuve une inscription (admin seulement)
 */
export function approveUser(userId: string): AuthResponse {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return { success: false, message: 'Action non autorisée' };
  }
  
  const db = getUsersDB();
  let targetEmail: string | null = null;
  
  // Trouver l'utilisateur par ID
  db.forEach((record, email) => {
    if (record.user.id === userId) {
      targetEmail = email;
    }
  });
  
  if (!targetEmail) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }
  
  const userRecord = db.get(targetEmail)!;
  userRecord.user.status = 'approved';
  userRecord.user.approvedBy = state.user!.email;
  userRecord.user.approvedAt = new Date().toISOString();
  
  db.set(targetEmail, userRecord);
  saveUsersDB(db);
  
  return {
    success: true,
    message: `Utilisateur ${userRecord.user.prenom} ${userRecord.user.nom} approuvé`,
    user: userRecord.user
  };
}

/**
 * Change le rôle d'un utilisateur (admin seulement)
 */
export function changeUserRole(userId: string, newRole: User['role']): AuthResponse {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return { success: false, message: 'Action non autorisée' };
  }

  // Empêcher de changer son propre rôle
  if (state.user?.id === userId) {
    return { success: false, message: 'Vous ne pouvez pas modifier votre propre rôle' };
  }

  const db = getUsersDB();
  let targetEmail: string | null = null;

  db.forEach((record, email) => {
    if (record.user.id === userId) {
      targetEmail = email;
    }
  });

  if (!targetEmail) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }

  const userRecord = db.get(targetEmail)!;
  const oldRole = userRecord.user.role;
  userRecord.user.role = newRole;

  db.set(targetEmail, userRecord);
  saveUsersDB(db);

  const roleLabel = newRole === 'admin' ? 'Administrateur' : newRole === 'viewer' ? 'Lecteur' : 'Utilisateur';
  return {
    success: true,
    message: `Rôle de ${userRecord.user.prenom} ${userRecord.user.nom} changé en ${roleLabel}`,
    user: userRecord.user
  };
}

/**
 * Rejette une inscription (admin seulement)
 */
export function rejectUser(userId: string): AuthResponse {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return { success: false, message: 'Action non autorisée' };
  }
  
  const db = getUsersDB();
  let targetEmail: string | null = null;
  
  // Trouver l'utilisateur par ID
  db.forEach((record, email) => {
    if (record.user.id === userId) {
      targetEmail = email;
    }
  });
  
  if (!targetEmail) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }
  
  const userRecord = db.get(targetEmail)!;
  userRecord.user.status = 'rejected';
  
  db.set(targetEmail, userRecord);
  saveUsersDB(db);
  
  return {
    success: true,
    message: `Inscription de ${userRecord.user.prenom} ${userRecord.user.nom} refusée`,
    user: userRecord.user
  };
}

/**
 * Supprime un utilisateur (admin seulement)
 */
export function deleteUser(userId: string): AuthResponse {
  const state = getAuthState();
  if (!state.isAuthenticated || state.user?.role !== 'admin') {
    return { success: false, message: 'Action non autorisée' };
  }
  
  // Empêcher la suppression de son propre compte
  if (state.user?.id === userId) {
    return { success: false, message: 'Vous ne pouvez pas supprimer votre propre compte' };
  }
  
  const db = getUsersDB();
  let targetEmail: string | null = null;
  let targetUser: User | null = null;
  
  // Trouver l'utilisateur par ID
  db.forEach((record, email) => {
    if (record.user.id === userId) {
      targetEmail = email;
      targetUser = record.user;
    }
  });
  
  if (!targetEmail || !targetUser) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }
  
  db.delete(targetEmail);
  saveUsersDB(db);
  
  return {
    success: true,
    message: `Utilisateur ${targetUser.prenom} ${targetUser.nom} supprimé`
  };
}

/**
 * Change le mot de passe
 */
export function changePassword(currentPassword: string, newPassword: string): AuthResponse {
  const state = getAuthState();
  
  if (!state.isAuthenticated || !state.user) {
    return { success: false, message: 'Non authentifié' };
  }
  
  if (newPassword.length < 8) {
    return { success: false, message: 'Le nouveau mot de passe doit contenir au moins 8 caractères' };
  }
  
  const db = getUsersDB();
  const userRecord = db.get(state.user.email);
  
  if (!userRecord) {
    return { success: false, message: 'Utilisateur non trouvé' };
  }
  
  // Vérifier l'ancien mot de passe
  if (userRecord.passwordHash !== hashPassword(currentPassword)) {
    return { success: false, message: 'Mot de passe actuel incorrect' };
  }
  
  // Mettre à jour le mot de passe
  db.set(state.user.email, {
    ...userRecord,
    passwordHash: hashPassword(newPassword)
  });
  saveUsersDB(db);
  
  return {
    success: true,
    message: 'Mot de passe modifié avec succès'
  };
}
