const { createClient } = require('@supabase/supabase-js');

// تێبینی: لێرەدا دەبێت URL و Key ی پڕۆژەکەی خۆت لە Supabase دابنێیت
const supabaseUrl = 'YOUR_SUPABASE_URL';
const supabaseKey = 'YOUR_SUPABASE_ANON_KEY';

const supabase = createClient(supabaseUrl, supabaseKey);

module.exports = supabase;
