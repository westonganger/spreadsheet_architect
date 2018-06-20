# Be sure to restart your server when you modify this file.

# Your secret key for verifying the integrity of signed cookies.
# If you change this key, all old signed cookies will become invalid!
# Make sure the secret is at least 30 characters and all random,
# no regular words or you'll be exposed to dictionary attacks.

gem_version = ActiveRecord.gem_version
if gem_version <= Gem::Version.new("5.1")
  Dummy::Application.config.secret_token = '4f337f0063fbb4a724dd8da15419679300da990ae4f6c94d36c714a3cd07e9653fc42d902cf33a9b9449a28e7eb2673f928172d65a090fa3c9156d6beea8d16c'
end
