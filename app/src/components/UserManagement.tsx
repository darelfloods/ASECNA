import React, { useState, useEffect, useCallback } from 'react';
import {
  User,
  getPendingUsers,
  getAllUsers,
  approveUser,
  rejectUser,
  deleteUser,
  changeUserRole
} from '../services/authService';

interface UserManagementProps {
  currentUser: User | null;
}

type ViewMode = 'pending' | 'all';

export const UserManagement: React.FC<UserManagementProps> = ({ currentUser }) => {
  const [viewMode, setViewMode] = useState<ViewMode>('pending');
  const [users, setUsers] = useState<User[]>([]);
  const [pendingCount, setPendingCount] = useState(0);
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [emailConfigured, setEmailConfigured] = useState<boolean | null>(null);

  useEffect(() => {
    fetch('http://localhost:3002/api/email-status')
      .then(r => r.json())
      .then(d => setEmailConfigured(d.configured))
      .catch(() => setEmailConfigured(false));
  }, []);

  const loadUsers = useCallback(() => {
    if (viewMode === 'pending') {
      setUsers(getPendingUsers());
    } else {
      setUsers(getAllUsers());
    }
    setPendingCount(getPendingUsers().length);
  }, [viewMode]);

  useEffect(() => {
    loadUsers();
  }, [loadUsers]);

  const showMessage = (type: 'success' | 'error', text: string) => {
    setMessage({ type, text });
    setTimeout(() => setMessage(null), 4000);
  };

  const handleApprove = async (userId: string) => {
    const userToApprove = users.find(u => u.id === userId);
    const result = approveUser(userId);
    if (result.success) {
      showMessage('success', result.message);
      loadUsers();

      // Envoyer l'email de notification à l'utilisateur
      if (userToApprove) {
        try {
          const response = await fetch('http://localhost:3002/api/send-approval-email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              email: userToApprove.email,
              prenom: userToApprove.prenom,
              nom: userToApprove.nom
            })
          });
          const data = await response.json();
          if (data.success) {
            showMessage('success', `${result.message} — Email envoyé à ${userToApprove.email}`);
          } else if (data.warning) {
            console.warn('Email non configuré:', data.warning);
          }
        } catch (emailErr) {
          console.warn('Impossible d\'envoyer l\'email de confirmation:', emailErr);
        }
      }
    } else {
      showMessage('error', result.message);
    }
  };

  const handleReject = (userId: string) => {
    if (!confirm('Êtes-vous sûr de vouloir refuser cette inscription ?')) return;
    const result = rejectUser(userId);
    if (result.success) {
      showMessage('success', result.message);
      loadUsers();
    } else {
      showMessage('error', result.message);
    }
  };

  const handleRoleChange = (userId: string, newRole: User['role']) => {
    if (!confirm(`Changer le rôle de cet utilisateur en "${newRole === 'admin' ? 'Administrateur' : newRole === 'viewer' ? 'Lecteur' : 'Utilisateur'}" ?`)) return;
    const result = changeUserRole(userId, newRole);
    if (result.success) {
      showMessage('success', result.message);
      loadUsers();
    } else {
      showMessage('error', result.message);
    }
  };

  const handleDelete = (userId: string) => {
    if (!confirm('Êtes-vous sûr de vouloir supprimer cet utilisateur ? Cette action est irréversible.')) return;
    const result = deleteUser(userId);
    if (result.success) {
      showMessage('success', result.message);
      loadUsers();
    } else {
      showMessage('error', result.message);
    }
  };

  const formatDate = (dateString: string) => {
    return new Date(dateString).toLocaleDateString('fr-FR', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const STATUS_TOOLTIPS: Record<User['status'], string> = {
    pending:  'Ce compte est en attente de validation.\nL\'utilisateur ne peut pas encore se connecter.',
    approved: 'Ce compte est actif.\nL\'utilisateur peut se connecter et utiliser l\'application.',
    rejected: 'Cette inscription a été refusée.\nL\'utilisateur ne peut pas se connecter.'
  };

  const ROLE_TOOLTIPS: Record<User['role'], string> = {
    admin:  'Administrateur — Accès total.\nPeut gérer les utilisateurs, approuver les inscriptions et accéder à toutes les fonctionnalités.',
    user:   'Utilisateur — Accès standard.\nPeut générer des documents (factures, missions, bons de commande) et consulter son historique.',
    viewer: 'Lecteur — Accès limité en lecture seule.\nPeut consulter l\'historique mais ne peut pas générer de documents.'
  };

  const getStatusBadge = (status: User['status']) => {
    const labels: Record<User['status'], string> = { pending: 'En attente', approved: 'Actif', rejected: 'Refusé' };
    return (
      <span className={`user-status-badge ${status} um-badge-hoverable`}>
        {labels[status]}
        <span className="um-tooltip">{STATUS_TOOLTIPS[status]}</span>
      </span>
    );
  };

  const getRoleBadge = (role: User['role']) => {
    const labels: Record<User['role'], string> = { admin: 'Admin', user: 'Utilisateur', viewer: 'Lecteur' };
    return (
      <span className={`user-role-badge ${role} um-badge-hoverable`}>
        {labels[role]}
        <span className="um-tooltip">{ROLE_TOOLTIPS[role]}</span>
      </span>
    );
  };

  // Vérifier si l'utilisateur est admin
  if (!currentUser || currentUser.role !== 'admin') {
    return (
      <div className="user-management-unauthorized">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="12" cy="12" r="10" />
          <line x1="12" y1="8" x2="12" y2="12" />
          <line x1="12" y1="16" x2="12.01" y2="16" />
        </svg>
        <h2>Accès non autorisé</h2>
        <p>Cette section est réservée aux administrateurs.</p>
      </div>
    );
  }

  return (
    <div className="user-management">
      {/* En-tête */}
      <div className="user-management-header">
        <h2>Gestion des utilisateurs</h2>
        <p>Gérez les inscriptions et les comptes utilisateurs</p>
      </div>

      {/* Bannière email non configuré */}
      {emailConfigured === false && (
        <div className="user-email-warning">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="18" height="18">
            <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
            <line x1="12" y1="9" x2="12" y2="13" />
            <line x1="12" y1="17" x2="12.01" y2="17" />
          </svg>
          <span>
            Les notifications par email sont désactivées. Éditez <strong>server/email-config.json</strong> pour activer l'envoi automatique d'emails lors de l'approbation d'un compte.
          </span>
        </div>
      )}

      {/* Message */}
      {message && (
        <div className={`user-message ${message.type}`}>
          {message.type === 'success' ? (
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
              <polyline points="22,4 12,14.01 9,11.01" />
            </svg>
          ) : (
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" />
              <line x1="15" y1="9" x2="9" y2="15" />
              <line x1="9" y1="9" x2="15" y2="15" />
            </svg>
          )}
          {message.text}
        </div>
      )}

      {/* Onglets */}
      <div className="user-management-tabs">
        <button
          className={`user-tab ${viewMode === 'pending' ? 'active' : ''}`}
          onClick={() => setViewMode('pending')}
        >
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10" />
            <polyline points="12,6 12,12 16,14" />
          </svg>
          Demandes en attente
          {pendingCount > 0 && <span className="user-tab-badge">{pendingCount}</span>}
        </button>
        <button
          className={`user-tab ${viewMode === 'all' ? 'active' : ''}`}
          onClick={() => setViewMode('all')}
        >
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
            <circle cx="9" cy="7" r="4" />
            <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
            <path d="M16 3.13a4 4 0 0 1 0 7.75" />
          </svg>
          Tous les utilisateurs
        </button>
      </div>

      {/* Liste des utilisateurs */}
      <div className="user-list">
        {users.length === 0 ? (
          <div className="user-list-empty">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
              <circle cx="9" cy="7" r="4" />
              <line x1="23" y1="11" x2="17" y2="11" />
            </svg>
            <p>
              {viewMode === 'pending' 
                ? 'Aucune demande d\'inscription en attente' 
                : 'Aucun utilisateur trouvé'}
            </p>
          </div>
        ) : (
          users.map((user) => (
            <div key={user.id} className="user-card">
              <div className="user-card-header">
                <div className="user-avatar">
                  {user.prenom.charAt(0)}{user.nom.charAt(0)}
                </div>
                <div className="user-info">
                  <h3>{user.prenom} {user.nom}</h3>
                  <p className="user-email">{user.email}</p>
                </div>
                <div className="user-badges">
                  {getStatusBadge(user.status)}
                  {getRoleBadge(user.role)}
                </div>
              </div>
              
              <div className="user-card-details">
                {user.matricule && (
                  <div className="user-detail">
                    <span className="label">Matricule:</span>
                    <span className="value">{user.matricule}</span>
                  </div>
                )}
                {user.service && (
                  <div className="user-detail">
                    <span className="label">Service:</span>
                    <span className="value">{user.service}</span>
                  </div>
                )}
                <div className="user-detail">
                  <span className="label">Inscrit le:</span>
                  <span className="value">{formatDate(user.createdAt)}</span>
                </div>
                {user.lastLogin && (
                  <div className="user-detail">
                    <span className="label">Dernière connexion:</span>
                    <span className="value">{formatDate(user.lastLogin)}</span>
                  </div>
                )}
              </div>

              {/* Actions */}
              <div className="user-card-actions">
                {user.status === 'pending' && (
                  <>
                    <button
                      className="user-action-btn approve"
                      onClick={() => handleApprove(user.id)}
                    >
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <polyline points="20,6 9,17 4,12" />
                      </svg>
                      Approuver
                    </button>
                    <button
                      className="user-action-btn reject"
                      onClick={() => handleReject(user.id)}
                    >
                      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <line x1="18" y1="6" x2="6" y2="18" />
                        <line x1="6" y1="6" x2="18" y2="18" />
                      </svg>
                      Refuser
                    </button>
                  </>
                )}

                {/* Sélecteur de rôle — tous les comptes sauf le sien */}
                {user.id !== currentUser?.id && user.status === 'approved' && (
                  <div className="user-role-select-wrapper">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14">
                      <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
                      <circle cx="9" cy="7" r="4" />
                      <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
                      <path d="M16 3.13a4 4 0 0 1 0 7.75" />
                    </svg>
                    <select
                      className="user-role-select"
                      value={user.role}
                      onChange={(e) => handleRoleChange(user.id, e.target.value as User['role'])}
                    >
                      <option value="user">Utilisateur</option>
                      <option value="admin">Administrateur</option>
                      <option value="viewer">Lecteur</option>
                    </select>
                  </div>
                )}

                {user.id !== currentUser?.id && user.role !== 'admin' && (
                  <button
                    className="user-action-btn delete"
                    onClick={() => handleDelete(user.id)}
                  >
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <polyline points="3,6 5,6 21,6" />
                      <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
                    </svg>
                    Supprimer
                  </button>
                )}
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
};

export default UserManagement;
