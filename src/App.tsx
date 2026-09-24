import PWABadge from "./PWABadge.tsx";
import { Play, Square } from "lucide-react";
import { useCallback, useEffect, useLayoutEffect, useState } from "react";
import ExcelJS from "exceljs";

interface ShiftDataItem {
  date: Date;
  time: {
    start: Date;
    end: Date;
  };
  tasks: string;
}

interface PrototypeShiftDataItem {
  date: string;
  time: {
    start: string;
    end: string;
  };
  tasks: string;
}

export default function App() {
  const lastBillingPhase = new Date();
  const [inShift, setIsShift] = useState(false);
  const [shiftData, setShiftData] = useState<ShiftDataItem[]>(() => {
    const shiftData = localStorage.getItem("all-shift-data");
    if (!shiftData) return [];

    const parsedData: PrototypeShiftDataItem[] = JSON.parse(shiftData);
    return parsedData.map((p): ShiftDataItem => ({
      date: new Date(p.date),
      time: {
        start: new Date(p.time.start),
        end: new Date(p.time.end),
      },
      tasks: p.tasks,
    }));
  });

  useLayoutEffect(() => {
    setIsShift("shift-start-time" in localStorage);
  }, []);

  useEffect(() => {
    localStorage.setItem(
      "all-shift-data",
      JSON.stringify(
        shiftData.map((s): PrototypeShiftDataItem => ({
          date: s.date.toISOString(),
          time: {
            start: s.time.start.toISOString(),
            end: s.time.end.toISOString(),
          },
          tasks: s.tasks,
        })),
      ),
    );
  }, [shiftData]);

  const ShiftIcon = inShift ? Square : Play;
  const toggleShift = useCallback(() => {
    if (!inShift) {
      localStorage.setItem("shift-start-time", Date.now().toString());
      setIsShift(true);
      return;
    }

    const start = new Date(
      parseInt(localStorage.getItem("shift-start-time") ?? "0"),
    );

    // Invalid time
    if (start.getTime() === 0) {
      console.error("Called stop shift when no shift is running");
      setIsShift(false);
      return;
    }

    const end = new Date();
    setShiftData((prev) => [
      ...prev,
      {
        date: new Date(),
        time: {
          start,
          end,
        },
        tasks: "",
      },
    ]);
    localStorage.removeItem("shift-start-time");
    setIsShift(false);
  }, [inShift]);

  const startNewBillingPhase = useCallback(() => {
    fetch("/timesheet.xlsx").then(async (res) => {
      const buffer = await res.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // First worksheet
      const worksheet = workbook.worksheets[0];

      const templateRowNumber = 11;

      if (shiftData.length === 0) {
        throw new Error("No shift data supplied.");
      }

      const templateRow = worksheet.getRow(templateRowNumber);
      const totalRowTemplate = worksheet.getRow(12);
      const totalRow = worksheet.getRow(11 + shiftData.length);

      // Copy the "total" row
      copyRow(totalRowTemplate, totalRow);

      // Update the "total" row's formula
      totalRow.getCell(7).value = {
        formula: `=SUM($G$11:$G$${10 + shiftData.length})`,
      };

      // Duplicate template row to fit all data
      for (let i = 1; i < shiftData.length; i++) {
        const o = templateRowNumber + i;
        const targetRow = worksheet.getRow(o);

        copyRow(templateRow, targetRow);

        // Update values accordingly
        targetRow.getCell(1).value = templateRowNumber + i - 10;
      }

      // Start filling in the values
      shiftData.forEach((shift, i) => {
        const o = templateRowNumber + i;
        const row = worksheet.getRow(o);

        row.getCell(2).value = toESTISO(shift.date);
        row.getCell(3).value = toESTISO(shift.time.start);
        row.getCell(4).value = toESTISO(shift.time.end);
        row.getCell(8).value = shift.tasks;
        row.getCell(7).value = {
          formula: `=IF(COUNTA(C${o}, D${o})=0, "", IF(D${o}<C${o},D${o}+12-C${o},D${o}-C${o}))`,
        };
      });

      // Add the start and end dates
      worksheet.getCell("F4").value = shiftData[0].date;
      worksheet.getCell("F5").value = shiftData.at(-1)!.date;

      // Generate the modified XLSX
      const output = await workbook.xlsx.writeBuffer();

      const blob = new Blob([output], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);

      const link = document.createElement("a");
      link.href = url;
      link.download = `WTN Timesheet (${shiftData[0].date.toLocaleDateString("en-US", { month: "short", day: "2-digit", timeZone: "America/New_York" })} - ${shiftData.at(-1)!.date.toLocaleDateString("en-US", { month: "short", day: "2-digit", timeZone: "America/New_York" })})`;
      link.click();

      URL.revokeObjectURL(url);

      function copyRow(sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row) {
        targetRow.height = sourceRow.height;

        for (let col = 1; col <= 10; col++) {
          copyCell(sourceRow.getCell(col), targetRow.getCell(col));
        }

        function copyCell(source: ExcelJS.Cell, target: ExcelJS.Cell) {
          // Copy the value/formula.
          if (
            source.value &&
            typeof source.value === "object" &&
            "formula" in source.value
          ) {
            target.value = {
              formula: source.value.formula,
              ...(source.value.result !== undefined
                ? { result: source.value.result }
                : {}),
            } as ExcelJS.CellFormulaValue;
          } else {
            target.value = source.value;
          }

          // ExcelJS style objects should be copied as a whole.
          target.style = source.style;

          // Preserve number formatting.
          target.numFmt = source.numFmt;

          if (source.note) {
            target.note = source.note;
          }
        }
      }
    });
  }, [shiftData]);

  return (
    <div className="fixed inset-0 flex p-12">
      <PWABadge />
      <aside className="min-w-sm">
        <h1 className="text-3xl font-bold">Punch</h1>
        <p className="text-gray-400">Punch your WTN IT shifts</p>

        <div className="py-8 flex flex-col gap-4 items-start">
          <button onClick={toggleShift} className={inShift ? "pulsing" : ""}>
            <ShiftIcon size={32} />
          </button>

          <button
            onClick={startNewBillingPhase}
            disabled={shiftData.length === 0}
          >
            {shiftData.length > 0
              ? "Save billing phase as xlsx"
              : "Start new billing phase by clocking in/out"}
          </button>
          <button
            onClick={() => setShiftData([])}
            disabled={shiftData.length === 0}
          >
            Clear Table
          </button>
        </div>
      </aside>
      <main className="flex-1 overflow-auto scrollbar-none">
        <h2 className="text-2xl font-semibold flex gap-4">
          Current Billing Phase
          <span className="text-gray-400">
            (since {toESTDate(lastBillingPhase)})
          </span>
        </h2>
        <table className="w-full mt-8 border-collapse">
          <thead>
            <tr className="*:border *:border-gray-700 *:py-2 bg-gray-900">
              <th>Date</th>
              <th>Time Start</th>
              <th>Time End</th>
              <th>Time Elapsed</th>
              <th>Tasks</th>
            </tr>
          </thead>
          <tbody>
            {shiftData.map((s, i) => (
              <tr
                className="*:border *:border-gray-700 *:px-2 *:py-4 *:text-center"
                key={`shift-item-${i}`}
              >
                <td>{toESTDate(s.date)}</td>
                <td>{toESTDate(s.time.start, { timeOnly: true })}</td>
                <td>{toESTDate(s.time.end, { timeOnly: true })}</td>
                <td>{calculateTimeElapsed(s.time.start, s.time.end)}</td>
                <td>
                  <textarea
                    className="min-w-0 outline-none"
                    defaultValue={s.tasks}
                    placeholder="Add task details here"
                    onChange={(e) => {
                      setShiftData((prev) => {
                        prev[i].tasks = e.target.value ?? s.tasks;
                        return [...prev];
                      });
                    }}
                  ></textarea>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </main>
    </div>
  );
}

function toESTDate(date: Date, { timeOnly = false } = {}): string {
  if (timeOnly)
    return date.toLocaleTimeString("en-US", {
      timeZone: "America/New_York",
      timeStyle: "short",
    });

  return date.toLocaleDateString("en-US", {
    timeZone: "America/New_York",
    dateStyle: "medium",
  });
}

const estFormatter = new Intl.DateTimeFormat('en-US', {
  timeZone: 'America/New_York',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  second: '2-digit'
});

function toESTISO(date: Date) {
  const parts = estFormatter.formatToParts(date);
  const year = parseInt(parts.find(p => p.type === 'year')?.value ?? '0');
  const month = parseInt(parts.find(p => p.type === 'month')?.value ?? '0') - 1;
  const day = parseInt(parts.find(p => p.type === 'day')?.value ?? '0');
  const hour = parseInt(parts.find(p => p.type === 'hour')?.value ?? '0');
  const minute = parseInt(parts.find(p => p.type === 'minute')?.value ?? '0');
  const second = parseInt(parts.find(p => p.type === 'second')?.value ?? '0');

  // Create a UTC date from the EST components
  return new Date(Date.UTC(year, month, day, hour, minute, second));
}

function calculateTimeElapsed(date1: Date, date2: Date): string {
  const ms = Math.abs(date2.getTime() - date1.getTime());
  const s = ms / 1000;
  const m = s / 60;
  const h = m / 60;

  return `${Math.floor(h) % 60}:${(Math.floor(m) % 60).toString().padStart(2, "0")}:${(Math.floor(s) % 60).toString().padStart(2, "0")}`;
}
