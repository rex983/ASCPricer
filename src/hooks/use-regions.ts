"use client";

import { useState, useEffect } from "react";
import type { Region } from "@/types/region";

/**
 * Fetch and cache available regions.
 */
export function useRegions() {
  const [regions, setRegions] = useState<Region[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function fetchRegions() {
      try {
        const res = await fetch("/api/pricing/regions");
        if (res.ok) {
          const data = await res.json();
          setRegions(data);
        }
      } catch (error) {
        console.error("Failed to fetch regions:", error);
      } finally {
        setLoading(false);
      }
    }

    fetchRegions();
  }, []);

  return { regions, loading };
}
