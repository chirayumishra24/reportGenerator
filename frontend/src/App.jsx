import AppProviders from './app/AppProviders.jsx';
import AppShell from './app/AppShell.jsx';

export default function App() {
  return (
    <AppProviders>
      <AppShell />
    </AppProviders>
  );
}
