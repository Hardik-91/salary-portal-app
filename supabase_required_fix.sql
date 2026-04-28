-- Run this once in Supabase SQL Editor to enable Block / Unblock Employee
ALTER TABLE public.employees
ADD COLUMN IF NOT EXISTS status text DEFAULT 'Active';

UPDATE public.employees
SET status = 'Active'
WHERE status IS NULL;
