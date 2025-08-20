module.exports = {
    apps: [
      {
        name: "backend",
        script: "./backend/app.py",
        interpreter: "./backend/venv/bin/python3"
      },
      {
        name: "frontend",
        script: "npm",
        args: "run serve",   // or: "run start" if dev, or "run build" + "pm2 serve build 3000"
        cwd: "./frontend"
      }
    ]
  }
  