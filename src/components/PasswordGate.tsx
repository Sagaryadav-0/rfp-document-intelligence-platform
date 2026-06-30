import { useEffect, useState, type FormEvent, type ReactNode } from "react";
import { Loader2, Eye, EyeOff } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from "sonner";

const TOKEN_KEY = "rfp-app-token";

export const PasswordGate = ({ children }: { children: ReactNode }) => {
  const [unlocked, setUnlocked] = useState(false);
  const [password, setPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const supabaseUrl = import.meta.env.VITE_SUPABASE_URL as string;
  const anonKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY as string;

  useEffect(() => {
    if (typeof window !== "undefined" && sessionStorage.getItem(TOKEN_KEY)) {
      setUnlocked(true);
    }
  }, []);

  const onSubmit = async (e: FormEvent) => {
    e.preventDefault();
    if (!password) return;
    setSubmitting(true);

    try {
      const response = await fetch(`${supabaseUrl}/functions/v1/verify-password`, {
        method: "POST",
        headers: {
          apikey: anonKey,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ password, type: "app" }),
      });

      const data = await response.json();
      if (!response.ok || !data.ok || !data.token) {
        toast.error("Incorrect password");
        return;
      }

      sessionStorage.setItem(TOKEN_KEY, data.token);
      setUnlocked(true);
    } catch (err) {
      console.warn("verify-password fetch failed:", err);
      toast.error("Could not verify password");
    } finally {
      setSubmitting(false);
    }
  };

  if (unlocked) return <>{children}</>;

  return (
    <main className="flex min-h-screen items-center justify-center bg-[image:var(--gradient-soft)] px-4">
      <Card className="w-full max-w-sm border-border/60 p-6 shadow-[var(--shadow-card)]">
        <div className="mb-5 flex flex-col items-center text-center">
          <div className="mb-3 flex h-14 w-14 items-center justify-center rounded-full bg-primary/10">
            <img src="/logo.svg" alt="Logo" className="h-8 w-8" />
          </div>
          <h1 className="text-lg font-semibold text-foreground">Protected access</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Enter the password to use this tool.
          </p>
        </div>
        <form onSubmit={onSubmit} className="space-y-3">
          <div className="space-y-1.5">
            <Label htmlFor="password">Password</Label>
            <div className="relative">
              <Input
                id="password"
                type={showPassword ? "text" : "password"}
                autoFocus
                value={password}
                onChange={(e) => setPassword(e.target.value)}
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
          <Button type="submit" className="w-full" disabled={submitting || !password}>
            {submitting ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}
            Unlock
          </Button>
        </form>
      </Card>
    </main>
  );
};
