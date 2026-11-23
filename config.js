// Shared settings for every page of the wedding site.
window.WEDDING = {
  // Google Apps Script web app URL (the RSVP Sheet backend). Blank = preview mode with sample data.
  endpoint: 'https://script.google.com/macros/s/AKfycbzRrszBy-tJLji3SZweFAogJ9NaUgCUUdRBYZwXHtq_23egM-T2a4FsQgRAJzXM6Baknw/exec',
  // Optional site password. Leave blank for no password. To turn one on, put the SHA-256 hash of the
  // lowercase password here (ask Claude to set it for you).
  passwordHash: ''
};
