import KenbanApp from "../kenban/KenbanApp";
import "../../kenban-utils/kenban.css";
import "../../kenban-utils/kenbanApp.css";

export function KenbanView() {
  return (
    <div className="flex h-full w-full overflow-hidden kenban-scope" style={{ position: 'absolute', inset: 0 }}>
      <KenbanApp />
    </div>
  );
}
