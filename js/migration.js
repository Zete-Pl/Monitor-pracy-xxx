const Migration = {
  /**
   * Sprawdza, czy stan wymaga migracji do wersji V3.
   * V3 charakteryzuje się podziałem na sekcje 'common' i 'months'.
   */
  isV2State: (state) => {
    if (!state) return false;
    // Jeśli stan posiada 'persons' i 'monthlySheets' na głównym poziomie, to jest to V2
    return Array.isArray(state.persons) && Array.isArray(state.monthlySheets);
  },

  /**
   * Konwertuje płaski stan V2 na strukturę V3.
   */
  v2ToV3: (v2State) => {
    console.log("MIGRACJA: Rozpoczynanie migracji bazy danych V2 -> V3...");
    
    // 1. Wyłuskanie danych wspólnych
    const common = {
      persons: v2State.persons || [],
      clients: v2State.clients || [],
      worksCatalog: v2State.worksCatalog || [],
      config: v2State.config || { taxRate: 0.055, zusFixedAmount: 1600.27 },
      settings: v2State.settings || {},
      version: 'v3'
    };

    // 2. Przygotowanie kontenerów na miesiące
    const months = {};

    // Funkcja pomocnicza do bezpiecznego przypisania do miesiąca
    const ensureMonth = (monthKey) => {
      if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) return null;
      if (!months[monthKey]) {
        months[monthKey] = {
          monthlySheets: [],
          worksSheets: [],
          expenses: [],
          monthSettings: v2State.monthSettings && v2State.monthSettings[monthKey] ? v2State.monthSettings[monthKey] : {}
        };
      }
      return months[monthKey];
    };

    // 3. Rozdzielenie arkuszy godzinowych (monthlySheets)
    if (Array.isArray(v2State.monthlySheets)) {
      v2State.monthlySheets.forEach(sheet => {
        const m = ensureMonth(sheet.month);
        if (m) m.monthlySheets.push(sheet);
      });
    }

    // 4. Rozdzielenie prac ryczałtowych (worksSheets)
    if (Array.isArray(v2State.worksSheets)) {
      v2State.worksSheets.forEach(sheet => {
        const m = ensureMonth(sheet.month);
        if (m) m.worksSheets.push(sheet);
      });
    }

    // 5. Rozdzielenie wydatków i zaliczek (expenses)
    // Zgodnie z planem: używamy daty wydatku, jeśli istnieje.
    if (Array.isArray(v2State.expenses)) {
      v2State.expenses.forEach(exp => {
        let monthKey = null;
        if (exp.date && exp.date.length >= 7) {
          monthKey = exp.date.substring(0, 7); // "YYYY-MM"
        }
        
        // Jeśli nie ma daty, szukamy czy jest przypisany do jakiegoś arkusza (legacy)
        // Jeśli nic nie znajdziemy, używamy aktualnie wybranego miesiąca z V2 (jako fallback)
        if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) {
          monthKey = v2State.selectedMonth || 'unknown';
        }

        const m = ensureMonth(monthKey);
        if (m) m.expenses.push(exp);
      });
    }

    // 6. Przeniesienie pozostałych ustawień miesięcy, które nie mają jeszcze danych
    if (v2State.monthSettings) {
      Object.keys(v2State.monthSettings).forEach(monthKey => {
        ensureMonth(monthKey);
      });
    }

    console.log(`MIGRACJA: Zakończono. Przetworzono ${Object.keys(months).length} miesięcy.`);
    
    return {
      common,
      months,
      selectedMonth: v2State.selectedMonth // Ten trafi do Local Settings w Store
    };
  }
};
