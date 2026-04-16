-- Replace all profiles with the current employee roster (67 users, exported 2026-04-14).
-- NOTE: profiles is shared with bbd-launcher. This wipe affects that app too.
--
-- FKs to profiles that block deletion (neither cascades):
--   asc_quotes.created_by, asc_customers.created_by
-- asc_sales_reps.profile_id is ON DELETE SET NULL, but we wipe it too for a clean slate.
-- asc_uploads.uploaded_by FK was dropped in 20260310155024.

BEGIN;

DELETE FROM asc_quotes;
DELETE FROM asc_customers;
DELETE FROM asc_sales_reps;
DELETE FROM profiles;

INSERT INTO profiles (email, full_name, role, office) VALUES
  ('adam@bigbuildingsdirect.com',           'Adam Niemann',          'sales_rep', 'Harbor'),
  ('a.narayan@bigbuildingsdirect.com',      'Adi Narayan',           'sales_rep', 'Harbor'),
  ('a.blust@bigbuildingsdirect.com',        'Aidan Blust',           'sales_rep', 'Harbor'),
  ('alex@bigbuildingsdirect.com',           'Alex Ambs',             'sales_rep', 'Harbor'),
  ('a.chase@bigbuildingsdirect.com',        'Alyssa Chase',          'sales_rep', 'Harbor'),
  ('andrew@bigbuildingsdirect.com',         'Andrew Wilson',         'sales_rep', 'Harbor'),
  ('angela@bigbuildingsdirect.com',         'Angela Allen',          'sales_rep', 'Harbor'),
  ('bill@bigbuildingsdirect.com',           'Bill Alexander',        'sales_rep', 'Harbor'),
  ('brandyn@bigbuildingsdirect.com',        'Brandyn Wnukowski',     'sales_rep', 'Harbor'),
  ('bates@bigbuildingsdirect.com',          'Brian Bates',           'sales_rep', 'Harbor'),
  ('b.harley@bigbuildingsdirect.com',       'Bryan Harley',          'sales_rep', 'Harbor'),
  ('projects@bigbuildingsdirect.com',       'Building Projects',     'sales_rep', 'Harbor'),
  ('cancellations@bigbuildingsdirect.com',  'Cancellations Department', 'sales_rep', 'Harbor'),
  ('c.wade@bigbuildingsdirect.com',         'Cayman Wade',           'sales_rep', 'Harbor'),
  ('c.murphy@bigbuildingsdirect.com',       'Chase Murphy',          'sales_rep', 'Harbor'),
  ('christy@bigbuildingsdirect.com',        'Christy C',             'sales_rep', 'Harbor'),
  ('customerorders@bigbuildingsdirect.com', 'Customer Order',        'sales_rep', 'Harbor'),
  ('customer@bigbuildingsdirect.com',       'Customer Customer',     'sales_rep', 'Harbor'),
  ('d.rodriguez@bigbuildingsdirect.com',    'Dariel Rodriguez',      'sales_rep', 'Harbor'),
  ('developer@bigbuildingsdirect.com',      'Developer K',           'sales_rep', 'Harbor'),
  ('devin@bigbuildingsdirect.com',          'Devin Cunningham',      'sales_rep', 'Harbor'),
  ('d.schmidt@bigbuildingsdirect.com',      'Dylan Schmidt',         'sales_rep', 'Harbor'),
  ('e.charnota@bigbuildingsdirect.com',     'Emily Charnota',        'sales_rep', 'Harbor'),
  ('e.quesada@bigbuildingsdirect.com',      'Emily Quesada',         'sales_rep', 'Harbor'),
  ('e.smeltzer@bigbuildingsdirect.com',     'Evan Smeltzer',         'sales_rep', 'Harbor'),
  ('gabe@bigbuildingsdirect.com',           'Gabriel De-Alba',       'sales_rep', 'Harbor'),
  ('garrett@bigbuildingsdirect.com',        'Garrett Ryder',         'sales_rep', 'Harbor'),
  ('jacob@bigbuildingsdirect.com',          'Jacob Reynolds',        'sales_rep', 'Harbor'),
  ('j.mayers@bigbuildingsdirect.com',       'Jakari Mayers',         'sales_rep', 'Harbor'),
  ('jason@bigbuildingsdirect.com',          'Jason Porcelli',        'sales_rep', 'Harbor'),
  ('j.lemon@bigbuildingsdirect.com',        'Jordan Lemon',          'sales_rep', 'Harbor'),
  ('jay@bigbuildingsdirect.com',            'Jordan Socarras',       'sales_rep', 'Harbor'),
  ('k.occe@bigbuildingsdirect.com',         'Kayani Occe',           'sales_rep', 'Harbor'),
  ('kelvin@bigbuildingsdirect.com',         'Kelvin Soto',           'sales_rep', 'Harbor'),
  ('l.arasimowicz@bigbuildingsdirect.com',  'Liliana Arasimowicz',   'sales_rep', 'Harbor'),
  ('macedona@bigbuildingsdirect.com',       'Macedona Big',          'sales_rep', 'Harbor'),
  ('manufacture@bigbuildingsdirect.com',    'Manufacture M',         'sales_rep', 'Harbor'),
  ('m.wright@bigbuildingsdirect.com',       'Max Wright',            'sales_rep', 'Harbor'),
  ('mayson@bigbuildingsdirect.com',         'Mayson Dunnigan',       'sales_rep', 'Harbor'),
  ('n.deboe@bigbuildingsdirect.com',        'Nicholas Deboe',        'sales_rep', 'Harbor'),
  ('nick@bigbuildingsdirect.com',           'Nick Brunsman',         'sales_rep', 'Harbor'),
  ('orders@bigbuildingsdirect.com',         'Order Processer',       'sales_rep', 'Harbor'),
  ('parker@bigbuildingsdirect.com',         'Parker Parzych',        'sales_rep', 'Harbor'),
  ('r.cavallo@bigbuildingsdirect.com',      'Ray Cavallo',           'sales_rep', 'Harbor'),
  ('reed@bigbuildingsdirect.com',           'Reed Hunt',             'sales_rep', 'Harbor'),
  ('revisions@bigbuildingsdirect.com',      'Revisions Department',  'sales_rep', 'Harbor'),
  ('rex@bigbuildingsdirect.com',            'Rex Wu',                'admin',     'Harbor'),
  ('richard@bigbuildingsdirect.com',        'Richard Kallay',        'sales_rep', 'Harbor'),
  ('r.lopez@bigbuildingsdirect.com',        'Rob Lopez',             'sales_rep', 'Harbor'),
  ('rob@bigbuildingsdirect.com',            'Rob Salaita',           'sales_rep', 'Harbor'),
  ('robin@bigbuildingsdirect.com',          'Robin Campbell',        'sales_rep', 'Harbor'),
  ('ryan@bigbuildingsdirect.com',           'Ryan Hamilton',         'sales_rep', 'Harbor'),
  ('sabrina@bigbuildingsdirect.com',        'Sabrina Big Buildings', 'sales_rep', 'Harbor'),
  ('sales@bigbuildingsdirect.com',          'Sales Big Buildings',   'sales_rep', 'Harbor'),
  ('salita@bigbuildingsdirect.com',         'Salita Bengochea',      'sales_rep', 'Harbor'),
  ('s.farabaugh@bigbuildingsdirect.com',    'Samantha Farabaugh',    'sales_rep', 'Harbor'),
  ('samantha@bigbuildingsdirect.com',       'Samantha Napoli',       'sales_rep', 'Harbor'),
  ('successteam@bigbuildingsdirect.com',    'Success Team',          'sales_rep', 'Harbor'),
  ('support@bigbuildingsdirect.com',        'Support Big Buildings', 'sales_rep', 'Harbor'),
  ('tim@bigbuildingsdirect.com',            'Tim Reynolds',          'sales_rep', 'Harbor'),
  ('timothy@bigbuildingsdirect.com',        'Timothy Hickman',       'sales_rep', 'Harbor'),
  ('t.woodmansee@bigbuildingsdirect.com',   'Tom Woodmansee',        'sales_rep', 'Harbor'),
  ('tony@bigbuildingsdirect.com',           'Tony Panapinto',        'sales_rep', 'Harbor'),
  ('tucker@bigbuildingsdirect.com',         'Tucker Fine',           'sales_rep', 'Harbor'),
  ('t.simpson@bigbuildingsdirect.com',      'Ty Simpson',            'sales_rep', 'Harbor'),
  ('t.hughes@bigbuildingsdirect.com',       'Tyler Hughes',          'sales_rep', 'Harbor'),
  ('y.pandit@bigbuildingsdirect.com',       'Yesha Pandit',          'sales_rep', 'Harbor');

COMMIT;
