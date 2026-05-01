export default function Home() {
  return (
    <main style={{ fontFamily: 'monospace', padding: '2rem' }}>
      <h1>Text to Voice API</h1>
      <ul>
        <li>GET /api/health</li>
        <li>POST /api/generate — body: {"{ tab_name, sheet_id?, date_filter? }"}</li>
        <li>POST /api/generate-all</li>
      </ul>
    </main>
  );
}
