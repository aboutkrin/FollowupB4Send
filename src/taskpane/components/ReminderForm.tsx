import * as React from "react";
import {
  Button,
  Field,
  Input,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Spinner,
  Title3,
} from "@fluentui/react-components";
import {
  CalendarLtr24Regular,
  Send24Regular,
} from "@fluentui/react-icons";
import { setFlagAndSend, sendWithoutFlag } from "../services/flagService";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "20px",
    maxWidth: "360px",
  },
  quickPicks: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  actions: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "8px",
  },
});

type QuickPick = "today" | "tomorrow" | "thisWeek" | "nextWeek" | "custom";

function getQuickPickDate(pick: QuickPick): { start: Date; due: Date } {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  switch (pick) {
    case "today":
      return { start: today, due: today };
    case "tomorrow": {
      const d = new Date(today);
      d.setDate(d.getDate() + 1);
      return { start: d, due: d };
    }
    case "thisWeek": {
      // Friday of the current week
      const d = new Date(today);
      const day = d.getDay();
      const diff = day <= 5 ? 5 - day : 0;
      d.setDate(d.getDate() + diff);
      return { start: today, due: d };
    }
    case "nextWeek": {
      // Next Monday → Next Friday
      const start = new Date(today);
      const day = start.getDay();
      const daysToMon = day === 0 ? 1 : 8 - day;
      start.setDate(start.getDate() + daysToMon);
      const due = new Date(start);
      due.setDate(due.getDate() + 4);
      return { start, due };
    }
    default:
      return { start: today, due: today };
  }
}

function toIso(date: Date): string {
  return date.toISOString();
}

function toInputDate(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export const ReminderForm: React.FC = () => {
  const styles = useStyles();
  const [selected, setSelected] = React.useState<QuickPick | null>(null);
  const [customDate, setCustomDate] = React.useState(toInputDate(new Date()));
  const [status, setStatus] = React.useState<
    { type: "idle" } | { type: "loading"; message: string } | { type: "success"; message: string } | { type: "error"; message: string }
  >({ type: "idle" });

  const handleSetReminder = async () => {
    if (!selected) return;

    let start: Date;
    let due: Date;

    if (selected === "custom") {
      const d = new Date(customDate + "T00:00:00");
      start = d;
      due = d;
    } else {
      const dates = getQuickPickDate(selected);
      start = dates.start;
      due = dates.due;
    }

    setStatus({ type: "loading", message: "Setting reminder and sending..." });
    try {
      await setFlagAndSend({
        startDate: toIso(start),
        dueDate: toIso(due),
      });
      setStatus({ type: "success", message: "Email sent with follow-up reminder!" });
    } catch (err: any) {
      if (err?.message === "SEND_UNAVAILABLE") {
        setStatus({
          type: "success",
          message: "Reminder set! Please click Send to send the email.",
        });
      } else {
        setStatus({ type: "error", message: err?.message ?? "Unknown error" });
      }
    }
  };

  const handleSendWithout = async () => {
    setStatus({ type: "loading", message: "Sending..." });
    try {
      await sendWithoutFlag();
      setStatus({ type: "success", message: "Email sent without reminder." });
    } catch (err: any) {
      if (err?.message === "SEND_UNAVAILABLE") {
        setStatus({
          type: "success",
          message: "Please click Send to send the email.",
        });
      } else {
        setStatus({ type: "error", message: err?.message ?? "Unknown error" });
      }
    }
  };

  const isLoading = status.type === "loading";

  return (
    <div className={styles.root}>
      <Title3>
        <CalendarLtr24Regular style={{ marginRight: 8, verticalAlign: "middle" }} />
        Set Follow-up Reminder
      </Title3>

      <div className={styles.quickPicks}>
        {(["today", "tomorrow", "thisWeek", "nextWeek", "custom"] as const).map(
          (pick) => (
            <Button
              key={pick}
              appearance={selected === pick ? "primary" : "secondary"}
              onClick={() => setSelected(pick)}
              disabled={isLoading}
            >
              {pick === "thisWeek"
                ? "This Week"
                : pick === "nextWeek"
                ? "Next Week"
                : pick.charAt(0).toUpperCase() + pick.slice(1)}
            </Button>
          )
        )}
      </div>

      {selected === "custom" && (
        <Field label="Reminder date">
          <Input
            type="date"
            value={customDate}
            onChange={(_, data) => setCustomDate(data.value)}
            disabled={isLoading}
          />
        </Field>
      )}

      <div className={styles.actions}>
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          onClick={handleSetReminder}
          disabled={!selected || isLoading}
        >
          Set Reminder &amp; Send
        </Button>
        <Button
          appearance="secondary"
          onClick={handleSendWithout}
          disabled={isLoading}
        >
          Send Without Reminder
        </Button>
      </div>

      {status.type === "loading" && <Spinner size="small" label={status.message} />}
      {status.type === "success" && (
        <MessageBar intent="success">
          <MessageBarBody>{status.message}</MessageBarBody>
        </MessageBar>
      )}
      {status.type === "error" && (
        <MessageBar intent="error">
          <MessageBarBody>{status.message}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
};
