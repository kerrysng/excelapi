module.exports = {
  apps: [{
    name: "Excel-Api",
    script: 'index1.js',
    instances: 3,
    log_date_format: 'DD-MM-YYYY HH:mm',
    error_file: 'err.log',
    out_file: 'out.log',
    log_file: 'combined.log',
    merge_logs:true,
    max_memory_restart: "1G",
    exec_mode: "cluster",
    autorestart: true,
    restart_delay: 5000,
    watch: false,

    env: {
      PORT: 3001,
      NODE_ENV: "development",
    },
    env_production: {
      PORT: 3000,
      NODE_ENV: "production",
    }
  }],

  deploy: {
    production: {
      user: 'SSH_USERNAME',
      host: 'SSH_HOSTMACHINE',
      ref: 'origin/master',
      repo: 'GIT_REPOSITORY',
      path: 'DESTINATION_PATH',
      'pre-deploy-local': '',
      'post-deploy': 'npm install && pm2 reload ecosystem.config.js --env production',
      'pre-setup': ''
    }
  }
};