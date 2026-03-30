// auth.js - 簡易認証（sessionStorage保持）

const Auth = {
  // ユーザー名・パスワードのSHA-256ハッシュ
  // デフォルト: admin / admin
  _credentialHash: '8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918:8c6976e5b5410415bde908bd4dee15dfb167a9c873fc4bb8a81f6f2ab448a918',

  async _sha256(text) {
    const data = new TextEncoder().encode(text);
    const hash = await crypto.subtle.digest('SHA-256', data);
    return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('');
  },

  async verify(username, password) {
    const userHash = await this._sha256(username);
    const passHash = await this._sha256(password);
    const combined = userHash + ':' + passHash;
    return combined === this._credentialHash;
  },

  isAuthenticated() {
    return sessionStorage.getItem('authenticated') === 'true';
  },

  setAuthenticated() {
    sessionStorage.setItem('authenticated', 'true');
  },

  /**
   * @param {Function} onSuccess - 認証成功時（既認証 or ログイン成功）に呼ばれるコールバック
   */
  init(onSuccess) {
    const overlay = document.getElementById('auth-overlay');
    if (this.isAuthenticated()) {
      overlay.classList.add('hidden');
      if (onSuccess) onSuccess();
      return;
    }

    const loginBtn = document.getElementById('auth-login-btn');
    const passInput = document.getElementById('auth-pass');
    const errorEl = document.getElementById('auth-error');

    const doLogin = async () => {
      const user = document.getElementById('auth-user').value.trim();
      const pass = passInput.value;
      errorEl.style.display = 'none';

      if (await this.verify(user, pass)) {
        this.setAuthenticated();
        overlay.classList.add('hidden');
        if (onSuccess) onSuccess();
      } else {
        errorEl.style.display = 'block';
        passInput.value = '';
        passInput.focus();
      }
    };

    loginBtn.addEventListener('click', doLogin);
    passInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') doLogin();
    });
  }
};
