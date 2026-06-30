import { useEffect, useState } from "react";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { toast } from "sonner";
import { Eye, EyeOff } from "lucide-react";

interface SessionData {
  id: string;
  session_id: string;
  file_name: string | null;
  file_size: number | null;
  status: string;
  metadata: any;
  created_at: string;
}

const TOKEN_KEY = "rfp-admin-token";

const Admin = () => {
  const [adminPassword, setAdminPassword] = useState("");
  const [isAdmin, setIsAdmin] = useState(false);
  const [sessions, setSessions] = useState<SessionData[]>([]);
  const [loading, setLoading] = useState(false);
  const [showPassword, setShowPassword] = useState(false);

  useEffect(() => {
    if (typeof window !== "undefined") {
      const token = sessionStorage.getItem(TOKEN_KEY);
      if (token) {
        setIsAdmin(true);
        fetchSessions();
      }
    }
  }, []);

  const fetchSessions = async () => {
    setLoading(true);
    try {
      const token = typeof window !== "undefined" ? sessionStorage.getItem(TOKEN_KEY) : null;
      if (!token) {
        throw new Error("Admin token missing. Please log in again.");
      }

      const response = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/admin-sessions`, {
        method: "POST",
        headers: {
          apikey: import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY as string,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ token }),
      });

      if (!response.ok) {
        throw new Error("Failed to load sessions");
      }

      const data = await response.json();
      setSessions(data.sessions || []);
    } catch (err) {
      toast.error("Failed to load sessions");
      console.error(err);
      setIsAdmin(false);
      sessionStorage.removeItem(TOKEN_KEY);
    } finally {
      setLoading(false);
    }
  };

  const handleAdminLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!adminPassword) {
      toast.error("Enter the admin password");
      return;
    }

    setLoading(true);
    try {
      const response = await fetch(`${import.meta.env.VITE_SUPABASE_URL}/functions/v1/verify-password`, {
        method: "POST",
        headers: {
          apikey: import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY as string,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ password: adminPassword, type: "admin" }),
      });

      if (!response.ok) {
        throw new Error("Incorrect admin password");
      }

      const data = await response.json();
      if (!data.ok || !data.token) {
        throw new Error("Incorrect admin password");
      }

      sessionStorage.setItem(TOKEN_KEY, data.token);
      setIsAdmin(true);
      setAdminPassword("");
      fetchSessions();
    } catch (err) {
      toast.error("Incorrect admin password");
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  if (!isAdmin) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-[image:var(--gradient-soft)] px-4">
        <Card className="w-full max-w-sm border-border/60 p-6 shadow-[var(--shadow-card)]">
          <div className="mb-5 flex flex-col items-center text-center">
            <div className="mb-3 flex h-14 w-14 items-center justify-center rounded-full bg-primary/10">
              <img src="/logo.svg" alt="Logo" className="h-8 w-8" />
            </div>
            <h1 className="text-lg font-semibold text-foreground">Admin Access</h1>
            <p className="mt-1 text-sm text-muted-foreground">
              Enter admin password to view dashboard.
            </p>
          </div>
          <form onSubmit={handleAdminLogin} className="space-y-3">
            <div className="space-y-1.5">
              <Label htmlFor="admin-password">Admin Password</Label>
              <div className="relative">
                <Input
                  id="admin-password"
                  type={showPassword ? "text" : "password"}
                  autoFocus
                  value={adminPassword}
                  onChange={(e) => setAdminPassword(e.target.value)}
                  placeholder="••••••••"
                  className="pr-10"
                />
                <Button
                  type="button"
                  variant="ghost"
                  size="sm"
                  className="absolute right-0 top-0 h-full px-3 py-2 hover:bg-transparent"
                  onClick={() => setShowPassword(!showPassword)}
                  tabIndex={-1}
                >
                  {showPassword ? (
                    <EyeOff className="h-4 w-4 text-muted-foreground" />
                  ) : (
                    <Eye className="h-4 w-4 text-muted-foreground" />
                  )}
                </Button>
              </div>
            </div>
            <Button type="submit" className="w-full" disabled={!adminPassword}>
              Access Admin
            </Button>
          </form>
        </Card>
      </div>
    );
  }

  return (
    <div className="container mx-auto px-4 py-8">
      <h1 className="text-2xl font-bold mb-6">Admin Dashboard</h1>
      
      <Card className="p-6">
        <h2 className="text-lg font-semibold mb-4">RFP Sessions</h2>
        {loading ? (
          <p>Loading...</p>
        ) : (
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Session ID</TableHead>
                <TableHead>File Name</TableHead>
                <TableHead>File Size</TableHead>
                <TableHead>Status</TableHead>
                <TableHead>Created At</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {sessions.map((session) => (
                <TableRow key={session.id}>
                  <TableCell className="font-mono text-sm">{session.session_id}</TableCell>
                  <TableCell>{session.file_name || "N/A"}</TableCell>
                  <TableCell>{session.file_size ? `${(session.file_size / 1024 / 1024).toFixed(2)} MB` : "N/A"}</TableCell>
                  <TableCell>{session.status}</TableCell>
                  <TableCell>{new Date(session.created_at).toLocaleString()}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        )}
        {sessions.length === 0 && !loading && (
          <p className="text-center text-muted-foreground mt-4">No sessions found.</p>
        )}
      </Card>
    </div>
  );
};

export default Admin;