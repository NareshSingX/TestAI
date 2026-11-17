const rules = [
  {
    category: "Action needed",
    terms: ["action required", "approval", "please review", "due today", "response needed"]
  },
  {
    category: "Meetings",
    terms: ["meeting", "invite", "calendar", "agenda", "minutes"]
  },
  {
    category: "Finance",
    terms: ["invoice", "receipt", "payment", "statement", "expense"]
  }
];

export function categorizeMessage(message) {
  const subject = typeof message.subject === "string" ? message.subject.toLowerCase() : "";

  for (const rule of rules) {
    if (rule.terms.some((term) => subject.includes(term))) {
      return rule.category;
    }
  }

  return "Other";
}

export function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => {
    const entities = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    };

    return entities[character];
  });
}
