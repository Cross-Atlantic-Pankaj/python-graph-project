module.exports = {
  apps: [
    {
      name: 'graph-project-backend',
      cwd: '/home/ubuntu/python-graph-project/backend',
      script: '/home/ubuntu/python-graph-project/venv/bin/gunicorn',
      args: '-c gunicorn.conf.py "app:create_app()"',
      interpreter: 'none',
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      watch: false,
      max_memory_restart: '1G',
      env: {
        NODE_ENV: 'production',
        PYTHONUNBUFFERED: '1'
      },
      error_file: '/home/ubuntu/.pm2/logs/backend-error.log',
      out_file: '/home/ubuntu/.pm2/logs/backend-out.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      merge_logs: true
    },
    {
      name: 'graph-project-frontend',
      cwd: '/home/ubuntu/python-graph-project/frontend-react',
      script: 'npm',
      args: 'start',
      interpreter: 'none',
      instances: 1,
      exec_mode: 'fork',
      autorestart: true,
      watch: false,
      max_memory_restart: '500M',
      env: {
        NODE_ENV: 'production',
        PORT: 3002
      },
      error_file: '/home/ubuntu/.pm2/logs/frontend-error.log',
      out_file: '/home/ubuntu/.pm2/logs/frontend-out.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      merge_logs: true
    }
  ]
};
