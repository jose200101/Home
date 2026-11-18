/** Endpoints anteriores: solo administrador, con sesión activa comprobada en servidor. */
function homeLegacyAuthorize_(args) {
  const envelope = args[args.length - 1];
  const session = homeSession_(
    envelope && typeof envelope === "object" ? envelope.homeToken : "",
  );
  if (!session.admin)
    throw new Error(
      "Los registros anteriores solo están disponibles para el administrador del hogar.",
    );
  return session;
}
