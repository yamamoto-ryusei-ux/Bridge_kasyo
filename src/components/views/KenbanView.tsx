import KenbanApp from "../kenban/KenbanApp";
import "../../kenban-utils/kenban.css";
import "../../kenban-utils/kenbanApp.css";

export function KenbanView() {
  return (
    <div className="flex h-full overflow-hidden kenban-scope">
      <KenbanApp />
    </div>
  );
}
