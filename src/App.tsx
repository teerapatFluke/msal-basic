import { useMemo, useState } from 'react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useIsAuthenticated, useMsal } from '@azure/msal-react';
import './App.css';

const scopes = ['User.Read'];

function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const account = useMemo(() => accounts[0] ?? null, [accounts]);
  const userId = useMemo(() => {
    if (!account) {
      return '';
    }

    const claims = account.idTokenClaims as Record<string, unknown> | undefined;
    const oid = claims?.oid;
    const sub = claims?.sub;

    if (typeof oid === 'string' && oid.length > 0) {
      return oid;
    }

    if (typeof sub === 'string' && sub.length > 0) {
      return sub;
    }

    return account.localAccountId || account.homeAccountId;
  }, [account]);

  const [token, setToken] = useState('');
  const [message, setMessage] = useState('ยังไม่ได้ล็อกอิน');

  const handleLogin = async () => {
    try {
      const loginResult = await instance.loginPopup({ scopes });
      const selectedAccount = loginResult.account ?? null;
      if (!selectedAccount) {
        setMessage('ล็อกอินสำเร็จ แต่ไม่พบบัญชีผู้ใช้');
        return;
      }

      const tokenResult = await instance.acquireTokenSilent({
        account: selectedAccount,
        scopes
      });

      setToken(tokenResult.accessToken);
      setMessage('ล็อกอินสำเร็จ และดึง token เรียบร้อย');
    } catch (error) {
      setMessage('ไม่สามารถล็อกอินได้');
      console.error(error);
    }
  };

  const handleGetToken = async () => {
    if (!account) {
      setMessage('ยังไม่พบบัญชี ให้ล็อกอินก่อน');
      return;
    }

    try {
      const tokenResult = await instance.acquireTokenSilent({
        account,
        scopes
      });
      setToken(tokenResult.accessToken);
      setMessage('รีเฟรช token สำเร็จ');
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        const tokenResult = await instance.acquireTokenPopup({
          account,
          scopes
        });
        setToken(tokenResult.accessToken);
        setMessage('ขอสิทธิ์เพิ่มและรับ token สำเร็จ');
        return;
      }

      setMessage('ดึง token ไม่สำเร็จ');
      console.error(error);
    }
  };

  const handleCopy = async () => {
    if (!token) {
      setMessage('ยังไม่มี token ให้คัดลอก');
      return;
    }

    try {
      await navigator.clipboard.writeText(token);
      setMessage('คัดลอก token แล้ว');
    } catch (error) {
      setMessage('คัดลอก token ไม่สำเร็จ');
      console.error(error);
    }
  };

  const handleLogout = async () => {
    await instance.logoutPopup();
    setToken('');
    setMessage('ออกจากระบบแล้ว');
  };

  return (
    <main className="page">
      <section className="card">
        <h1>Azure Login Token Viewer</h1>
        <p className="status">สถานะ: {message}</p>

        <div className="actions">
          {!isAuthenticated && (
            <button onClick={handleLogin} className="primary">Login with Azure</button>
          )}
          {isAuthenticated && (
            <>
              <button onClick={handleGetToken} className="primary">Get Access Token</button>
              <button onClick={handleLogout} className="ghost">Logout</button>
            </>
          )}
        </div>

        {account && <p className="account">ผู้ใช้: {account.username}</p>}
        {account && <p className="account">User ID: {userId}</p>}

        <label htmlFor="token">Access Token</label>
        <textarea
          id="token"
          value={token}
          readOnly
          placeholder="Token จะแสดงที่นี่หลังจากล็อกอิน"
          rows={10}
        />

        <button onClick={handleCopy} className="copy">Copy Token</button>
      </section>
    </main>
  );
}

export default App;
