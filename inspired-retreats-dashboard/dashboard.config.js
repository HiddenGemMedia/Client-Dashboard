window.DASHBOARD_CONFIG = {
  supabaseUrl: "https://wdntkxamhjmrapifhhws.supabase.co",
  supabaseAnonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndkbnRreGFtaGptcmFwaWZoaHdzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYyNjYzNDcsImV4cCI6MjA5MTg0MjM0N30.tKejcUx7tS4sNgsp1HBUnAbmF_D63CloQ11vU_t7Y6c",
  tables: {
    clients: "clients",
    monthlyAds: "monthly_ad_data",
    roiMonthly: "roi_monthly"
  },
  defaults: {
    month: "",
    client: "inspired-retreats",
    from: "2026-01",
    to: "2026-03"
  },
  campaignOrder: ["Followers", "New Leads", "Retargeting"],
  benchmarks: {
    roas: { decent: 3, solid: 5, great: 10 },
    costPerVisit: { decent: 0.3, solid: 0.2, great: 0.1 },
    costPerLeadFollower: { decent: 0.8, solid: 0.5, great: 0.3 },
    percentOfBookingValue: { decent: 40, solid: 25, great: 15 }
  }
};
