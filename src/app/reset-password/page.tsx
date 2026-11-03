"use client";

import { FormEvent, useEffect, useState } from "react";
import { supabase } from "../lib/supabase";

export default function ResetPasswordPage() {
  const [newPassword, setNewPassword] = useState("");
  const [message, setMessage] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isReady, setIsReady] = useState(false);

  useEffect(() => {
    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") {
        setIsReady(true);
      }
    });

    async function checkRecoverySession() {
      const { data } = await supabase.auth.getSession();

      if (data.session) {
        setIsReady(true);
      }
    }

    checkRecoverySession();

    return () => subscription.unsubscribe();
  }, []);

  async function handlePasswordUpdate(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    setMessage("");
    setIsLoading(true);

    const { error } = await supabase.auth.updateUser({
      password: newPassword,
    });

    setIsLoading(false);

    if (error) {
      setMessage("Não foi possível salvar a senha. Solicite um novo link.");
      return;
    }

    setMessage(
      "Senha atualizada com sucesso. Volte para a tela de login e entre com a nova senha."
    );
    setNewPassword("");
  }

  return (
    <main className="min-h-screen bg-slate-100 px-4 py-10">
      <div className="mx-auto max-w-md rounded-2xl bg-white p-8 shadow-lg">
        <p className="text-sm font-semibold uppercase tracking-wider text-emerald-700">
          Sistema Comercial
        </p>

        <h1 className="mt-2 text-3xl font-bold text-slate-900">
          Criar nova senha
        </h1>

        <p className="mt-3 text-slate-600">
          Escolha uma nova senha para o acesso do Gestor Comercial.
        </p>

        {isReady ? (
          <form className="mt-8 space-y-5" onSubmit={handlePasswordUpdate}>
            <div>
              <label
                htmlFor="new-password"
                className="mb-2 block text-sm font-medium text-slate-700"
              >
                Nova senha
              </label>

              <input
                id="new-password"
                type="password"
                required
                minLength={6}
                value={newPassword}
                onChange={(event) => setNewPassword(event.target.value)}
                placeholder="Digite a nova senha"
                className="w-full rounded-lg border border-slate-300 px-4 py-3 text-slate-900 outline-none focus:border-emerald-600 focus:ring-2 focus:ring-emerald-200"
              />
            </div>

            <button
              type="submit"
              disabled={isLoading}
              className="w-full rounded-lg bg-emerald-700 px-4 py-3 font-semibold text-white transition hover:bg-emerald-800 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {isLoading ? "Salvando..." : "Salvar nova senha"}
            </button>
          </form>
        ) : (
          <p className="mt-6 rounded-lg bg-amber-50 p-3 text-sm text-amber-800">
            Abra esta página pelo link de recuperação de senha enviado pelo
            Supabase.
          </p>
        )}

        {message && (
          <p className="mt-5 rounded-lg bg-slate-100 p-3 text-sm text-slate-700">
            {message}
          </p>
        )}
      </div>
    </main>
  );
}