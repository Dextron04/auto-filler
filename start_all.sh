#!/bin/bash

# Auto-Filler Start Script

# Get the directory of this script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

# Function to kill background processes on exit
cleanup() {
    echo "Stopping servers..."
    kill $(jobs -p)
    exit
}

trap cleanup SIGINT SIGTERM

echo "--- Starting Auto-Filler ---"

# Start Backend (Flask)
echo "Starting Backend API on port 5000..."
cd "$DIR/backend"
../.venv/bin/python3 app.py &

# Start Frontend (Vite)
echo "Starting Frontend on port 5173..."
cd "$DIR/frontend"
npm run dev -- --host &

echo "--- Servers are running! ---"
echo "Access the app at http://localhost:5173"
echo "Press Ctrl+C to stop both servers."

wait
